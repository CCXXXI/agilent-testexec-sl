Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading

Friend Class frmMain
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
    Public WithEvents tmrFailure As System.Windows.Forms.Timer
    Public WithEvents chkReportPassedTests As System.Windows.Forms.CheckBox
    Public WithEvents chkReportFailedTests As System.Windows.Forms.CheckBox
    Public WithEvents chkReportException As System.Windows.Forms.CheckBox
    Public WithEvents chkShowReport As System.Windows.Forms.CheckBox
    Public WithEvents fraTestplanConfiguration As System.Windows.Forms.GroupBox
    Public WithEvents cmdSelectVariant As System.Windows.Forms.Button
    Public WithEvents cmdLoadTestplan As System.Windows.Forms.Button
    Public WithEvents cmdTxSLExit As System.Windows.Forms.Button
    Public WithEvents cmdLogin As System.Windows.Forms.Button
    Public WithEvents medBarCode As AxMSMask.AxMaskEdBox
    Public WithEvents fraTxSLConfiguration As System.Windows.Forms.GroupBox
    Public WithEvents fraReport As System.Windows.Forms.GroupBox
    Public WithEvents fraOperatorMessage As System.Windows.Forms.GroupBox
    Public WithEvents fraTestplanProgress As System.Windows.Forms.GroupBox
    Public WithEvents txtSerialNumber As System.Windows.Forms.TextBox
    Public WithEvents cmdRun As System.Windows.Forms.Button
    Public WithEvents cmdStop As System.Windows.Forms.Button
    Public WithEvents cmdAbort As System.Windows.Forms.Button
    Public WithEvents lblSerialNumber As System.Windows.Forms.Label
    Public WithEvents fraExecution As System.Windows.Forms.GroupBox
    Public WithEvents fraSystemStatus As System.Windows.Forms.GroupBox
    Public WithEvents imglogo As System.Windows.Forms.PictureBox
    Public WithEvents TestExecSL1 As AxHPTestExecSL.AxTestExecSL
    Public WithEvents lblTitle As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents rtbViewErrors As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbOperatorMessage As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbSystemStatus As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbReport As System.Windows.Forms.RichTextBox
    Friend WithEvents prbTestplan As System.Windows.Forms.ProgressBar
    Friend WithEvents staDescriptions As System.Windows.Forms.StatusBar
    Friend WithEvents staStats As System.Windows.Forms.StatusBar
    Friend WithEvents sbpPassed As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpFailed As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpTotal As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpYield As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpSince As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpUutName As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpTestplanName As System.Windows.Forms.StatusBarPanel
    Friend WithEvents sbpVariant As System.Windows.Forms.StatusBarPanel
    Friend WithEvents txtExecutionMode As System.Windows.Forms.Label
    Friend WithEvents lblExecutionMode As System.Windows.Forms.Label
    Public WithEvents lblBarCode As System.Windows.Forms.Label
    Friend WithEvents SysBox As System.Windows.Forms.PictureBox
    Friend WithEvents SysLabel As System.Windows.Forms.Label
    Friend WithEvents FixLabel As System.Windows.Forms.Label
    Friend WithEvents FixBox As System.Windows.Forms.PictureBox
    Friend WithEvents DebugCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents DebugNumericUpDown As System.Windows.Forms.NumericUpDown
    Friend WithEvents sbpCurrentTestName As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkReportPassedTests = New System.Windows.Forms.CheckBox()
        Me.chkReportFailedTests = New System.Windows.Forms.CheckBox()
        Me.chkReportException = New System.Windows.Forms.CheckBox()
        Me.chkShowReport = New System.Windows.Forms.CheckBox()
        Me.cmdSelectVariant = New System.Windows.Forms.Button()
        Me.cmdLoadTestplan = New System.Windows.Forms.Button()
        Me.cmdTxSLExit = New System.Windows.Forms.Button()
        Me.cmdLogin = New System.Windows.Forms.Button()
        Me.txtSerialNumber = New System.Windows.Forms.TextBox()
        Me.cmdRun = New System.Windows.Forms.Button()
        Me.cmdStop = New System.Windows.Forms.Button()
        Me.cmdAbort = New System.Windows.Forms.Button()
        Me.imglogo = New System.Windows.Forms.PictureBox()
        Me.tmrFailure = New System.Windows.Forms.Timer(Me.components)
        Me.fraTestplanConfiguration = New System.Windows.Forms.GroupBox()
        Me.fraTxSLConfiguration = New System.Windows.Forms.GroupBox()
        Me.medBarCode = New AxMSMask.AxMaskEdBox()
        Me.lblBarCode = New System.Windows.Forms.Label()
        Me.fraReport = New System.Windows.Forms.GroupBox()
        Me.rtbReport = New System.Windows.Forms.RichTextBox()
        Me.fraOperatorMessage = New System.Windows.Forms.GroupBox()
        Me.rtbViewErrors = New System.Windows.Forms.RichTextBox()
        Me.rtbOperatorMessage = New System.Windows.Forms.RichTextBox()
        Me.fraTestplanProgress = New System.Windows.Forms.GroupBox()
        Me.prbTestplan = New System.Windows.Forms.ProgressBar()
        Me.fraExecution = New System.Windows.Forms.GroupBox()
        Me.txtExecutionMode = New System.Windows.Forms.Label()
        Me.lblExecutionMode = New System.Windows.Forms.Label()
        Me.lblSerialNumber = New System.Windows.Forms.Label()
        Me.fraSystemStatus = New System.Windows.Forms.GroupBox()
        Me.rtbSystemStatus = New System.Windows.Forms.RichTextBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.staDescriptions = New System.Windows.Forms.StatusBar()
        Me.sbpUutName = New System.Windows.Forms.StatusBarPanel()
        Me.sbpTestplanName = New System.Windows.Forms.StatusBarPanel()
        Me.sbpVariant = New System.Windows.Forms.StatusBarPanel()
        Me.sbpCurrentTestName = New System.Windows.Forms.StatusBarPanel()
        Me.staStats = New System.Windows.Forms.StatusBar()
        Me.sbpPassed = New System.Windows.Forms.StatusBarPanel()
        Me.sbpFailed = New System.Windows.Forms.StatusBarPanel()
        Me.sbpTotal = New System.Windows.Forms.StatusBarPanel()
        Me.sbpYield = New System.Windows.Forms.StatusBarPanel()
        Me.sbpSince = New System.Windows.Forms.StatusBarPanel()
        Me.TestExecSL1 = New AxHPTestExecSL.AxTestExecSL()
        Me.SysBox = New System.Windows.Forms.PictureBox()
        Me.SysLabel = New System.Windows.Forms.Label()
        Me.FixLabel = New System.Windows.Forms.Label()
        Me.FixBox = New System.Windows.Forms.PictureBox()
        Me.DebugCheckBox = New System.Windows.Forms.CheckBox()
        Me.DebugNumericUpDown = New System.Windows.Forms.NumericUpDown()
        CType(Me.imglogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraTestplanConfiguration.SuspendLayout()
        Me.fraTxSLConfiguration.SuspendLayout()
        CType(Me.medBarCode, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraReport.SuspendLayout()
        Me.fraOperatorMessage.SuspendLayout()
        Me.fraTestplanProgress.SuspendLayout()
        Me.fraExecution.SuspendLayout()
        Me.fraSystemStatus.SuspendLayout()
        CType(Me.sbpUutName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpTestplanName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpVariant, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpCurrentTestName, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpPassed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpFailed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpTotal, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpYield, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.sbpSince, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TestExecSL1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SysBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FixBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DebugNumericUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'chkReportPassedTests
        '
        Me.chkReportPassedTests.BackColor = System.Drawing.SystemColors.Control
        Me.chkReportPassedTests.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReportPassedTests.Enabled = False
        Me.chkReportPassedTests.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkReportPassedTests.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkReportPassedTests.Location = New System.Drawing.Point(14, 18)
        Me.chkReportPassedTests.Name = "chkReportPassedTests"
        Me.chkReportPassedTests.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReportPassedTests.Size = New System.Drawing.Size(189, 20)
        Me.chkReportPassedTests.TabIndex = 9
        Me.chkReportPassedTests.Text = "Report &Passed Tests"
        Me.ToolTip1.SetToolTip(Me.chkReportPassedTests, "Show passing tests in the report.")
        Me.chkReportPassedTests.UseVisualStyleBackColor = False
        '
        'chkReportFailedTests
        '
        Me.chkReportFailedTests.BackColor = System.Drawing.SystemColors.Control
        Me.chkReportFailedTests.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReportFailedTests.Enabled = False
        Me.chkReportFailedTests.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkReportFailedTests.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkReportFailedTests.Location = New System.Drawing.Point(213, 18)
        Me.chkReportFailedTests.Name = "chkReportFailedTests"
        Me.chkReportFailedTests.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReportFailedTests.Size = New System.Drawing.Size(139, 20)
        Me.chkReportFailedTests.TabIndex = 10
        Me.chkReportFailedTests.Text = "Report &Failed Tests"
        Me.ToolTip1.SetToolTip(Me.chkReportFailedTests, "Show failing tests in the report.")
        Me.chkReportFailedTests.UseVisualStyleBackColor = False
        '
        'chkReportException
        '
        Me.chkReportException.BackColor = System.Drawing.SystemColors.Control
        Me.chkReportException.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkReportException.Enabled = False
        Me.chkReportException.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkReportException.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkReportException.Location = New System.Drawing.Point(14, 41)
        Me.chkReportException.Name = "chkReportException"
        Me.chkReportException.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkReportException.Size = New System.Drawing.Size(188, 20)
        Me.chkReportException.TabIndex = 11
        Me.chkReportException.Text = "Report &Exceptions"
        Me.ToolTip1.SetToolTip(Me.chkReportException, "Show unhandled exceptions in the report.")
        Me.chkReportException.UseVisualStyleBackColor = False
        '
        'chkShowReport
        '
        Me.chkShowReport.BackColor = System.Drawing.SystemColors.Control
        Me.chkShowReport.Checked = True
        Me.chkShowReport.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowReport.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkShowReport.Enabled = False
        Me.chkShowReport.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkShowReport.Location = New System.Drawing.Point(213, 41)
        Me.chkShowReport.Name = "chkShowReport"
        Me.chkShowReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkShowReport.Size = New System.Drawing.Size(139, 20)
        Me.chkShowReport.TabIndex = 12
        Me.chkShowReport.Text = "Sho&w Report"
        Me.ToolTip1.SetToolTip(Me.chkShowReport, "Show the report.")
        Me.chkShowReport.UseVisualStyleBackColor = False
        '
        'cmdSelectVariant
        '
        Me.cmdSelectVariant.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSelectVariant.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSelectVariant.Enabled = False
        Me.cmdSelectVariant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSelectVariant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSelectVariant.Location = New System.Drawing.Point(189, 25)
        Me.cmdSelectVariant.Name = "cmdSelectVariant"
        Me.cmdSelectVariant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSelectVariant.Size = New System.Drawing.Size(163, 63)
        Me.cmdSelectVariant.TabIndex = 2
        Me.cmdSelectVariant.Text = "Select &Variant..."
        Me.ToolTip1.SetToolTip(Me.cmdSelectVariant, "Selects a new testplan variant.")
        Me.cmdSelectVariant.UseVisualStyleBackColor = False
        '
        'cmdLoadTestplan
        '
        Me.cmdLoadTestplan.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLoadTestplan.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLoadTestplan.Enabled = False
        Me.cmdLoadTestplan.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLoadTestplan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLoadTestplan.Location = New System.Drawing.Point(8, 25)
        Me.cmdLoadTestplan.Name = "cmdLoadTestplan"
        Me.cmdLoadTestplan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLoadTestplan.Size = New System.Drawing.Size(161, 63)
        Me.cmdLoadTestplan.TabIndex = 1
        Me.cmdLoadTestplan.Text = "&Load Testplan..."
        Me.ToolTip1.SetToolTip(Me.cmdLoadTestplan, "Loads a new testplan")
        Me.cmdLoadTestplan.UseVisualStyleBackColor = False
        '
        'cmdTxSLExit
        '
        Me.cmdTxSLExit.BackColor = System.Drawing.SystemColors.Control
        Me.cmdTxSLExit.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdTxSLExit.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTxSLExit.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdTxSLExit.Location = New System.Drawing.Point(279, 25)
        Me.cmdTxSLExit.Name = "cmdTxSLExit"
        Me.cmdTxSLExit.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdTxSLExit.Size = New System.Drawing.Size(73, 33)
        Me.cmdTxSLExit.TabIndex = 3
        Me.cmdTxSLExit.Text = "E&xit"
        Me.ToolTip1.SetToolTip(Me.cmdTxSLExit, "Exits the application.")
        Me.cmdTxSLExit.UseVisualStyleBackColor = False
        Me.cmdTxSLExit.Visible = False
        '
        'cmdLogin
        '
        Me.cmdLogin.BackColor = System.Drawing.SystemColors.Control
        Me.cmdLogin.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdLogin.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdLogin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdLogin.Location = New System.Drawing.Point(8, 25)
        Me.cmdLogin.Name = "cmdLogin"
        Me.cmdLogin.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdLogin.Size = New System.Drawing.Size(73, 33)
        Me.cmdLogin.TabIndex = 0
        Me.cmdLogin.Text = "Lo&gin"
        Me.ToolTip1.SetToolTip(Me.cmdLogin, "Login a new operator")
        Me.cmdLogin.UseVisualStyleBackColor = False
        Me.cmdLogin.Visible = False
        '
        'txtSerialNumber
        '
        Me.txtSerialNumber.AcceptsReturn = True
        Me.txtSerialNumber.BackColor = System.Drawing.SystemColors.Window
        Me.txtSerialNumber.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSerialNumber.Enabled = False
        Me.txtSerialNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSerialNumber.Location = New System.Drawing.Point(152, 17)
        Me.txtSerialNumber.MaxLength = 0
        Me.txtSerialNumber.Name = "txtSerialNumber"
        Me.txtSerialNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSerialNumber.Size = New System.Drawing.Size(181, 20)
        Me.txtSerialNumber.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txtSerialNumber, "The serial number of the UUT is entered, and shown, here.")
        '
        'cmdRun
        '
        Me.cmdRun.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRun.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRun.Enabled = False
        Me.cmdRun.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRun.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRun.Location = New System.Drawing.Point(24, 62)
        Me.cmdRun.Name = "cmdRun"
        Me.cmdRun.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRun.Size = New System.Drawing.Size(73, 49)
        Me.cmdRun.TabIndex = 6
        Me.cmdRun.Text = "&Run"
        Me.ToolTip1.SetToolTip(Me.cmdRun, "Runs or continues the testplan")
        Me.cmdRun.UseVisualStyleBackColor = False
        '
        'cmdStop
        '
        Me.cmdStop.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStop.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStop.Enabled = False
        Me.cmdStop.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStop.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStop.Location = New System.Drawing.Point(137, 62)
        Me.cmdStop.Name = "cmdStop"
        Me.cmdStop.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStop.Size = New System.Drawing.Size(73, 49)
        Me.cmdStop.TabIndex = 7
        Me.cmdStop.Text = "&Stop"
        Me.ToolTip1.SetToolTip(Me.cmdStop, "Stops the testplan at the first opportunity, and executes clean up actions")
        Me.cmdStop.UseVisualStyleBackColor = False
        '
        'cmdAbort
        '
        Me.cmdAbort.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAbort.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAbort.Enabled = False
        Me.cmdAbort.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAbort.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAbort.Location = New System.Drawing.Point(248, 62)
        Me.cmdAbort.Name = "cmdAbort"
        Me.cmdAbort.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAbort.Size = New System.Drawing.Size(73, 49)
        Me.cmdAbort.TabIndex = 8
        Me.cmdAbort.Text = "&Abort"
        Me.ToolTip1.SetToolTip(Me.cmdAbort, "Aborts the tesplan immediately, and does NOT perform cleanup actions")
        Me.cmdAbort.UseVisualStyleBackColor = False
        '
        'imglogo
        '
        Me.imglogo.Cursor = System.Windows.Forms.Cursors.Default
        Me.imglogo.Image = CType(resources.GetObject("imglogo.Image"), System.Drawing.Image)
        Me.imglogo.Location = New System.Drawing.Point(16, 8)
        Me.imglogo.Name = "imglogo"
        Me.imglogo.Size = New System.Drawing.Size(36, 36)
        Me.imglogo.TabIndex = 30
        Me.imglogo.TabStop = False
        Me.ToolTip1.SetToolTip(Me.imglogo, "Replace this image with one of your chosing")
        '
        'tmrFailure
        '
        Me.tmrFailure.Interval = 60000
        '
        'fraTestplanConfiguration
        '
        Me.fraTestplanConfiguration.BackColor = System.Drawing.SystemColors.Control
        Me.fraTestplanConfiguration.Controls.Add(Me.chkReportPassedTests)
        Me.fraTestplanConfiguration.Controls.Add(Me.chkReportFailedTests)
        Me.fraTestplanConfiguration.Controls.Add(Me.chkReportException)
        Me.fraTestplanConfiguration.Controls.Add(Me.chkShowReport)
        Me.fraTestplanConfiguration.Enabled = False
        Me.fraTestplanConfiguration.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraTestplanConfiguration.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTestplanConfiguration.Location = New System.Drawing.Point(11, 325)
        Me.fraTestplanConfiguration.Name = "fraTestplanConfiguration"
        Me.fraTestplanConfiguration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTestplanConfiguration.Size = New System.Drawing.Size(360, 69)
        Me.fraTestplanConfiguration.TabIndex = 22
        Me.fraTestplanConfiguration.TabStop = False
        Me.fraTestplanConfiguration.Text = "Testplan Configuration"
        '
        'fraTxSLConfiguration
        '
        Me.fraTxSLConfiguration.BackColor = System.Drawing.SystemColors.Control
        Me.fraTxSLConfiguration.Controls.Add(Me.cmdSelectVariant)
        Me.fraTxSLConfiguration.Controls.Add(Me.cmdLoadTestplan)
        Me.fraTxSLConfiguration.Controls.Add(Me.cmdTxSLExit)
        Me.fraTxSLConfiguration.Controls.Add(Me.cmdLogin)
        Me.fraTxSLConfiguration.Controls.Add(Me.medBarCode)
        Me.fraTxSLConfiguration.Controls.Add(Me.lblBarCode)
        Me.fraTxSLConfiguration.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraTxSLConfiguration.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTxSLConfiguration.Location = New System.Drawing.Point(11, 400)
        Me.fraTxSLConfiguration.Name = "fraTxSLConfiguration"
        Me.fraTxSLConfiguration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTxSLConfiguration.Size = New System.Drawing.Size(361, 97)
        Me.fraTxSLConfiguration.TabIndex = 20
        Me.fraTxSLConfiguration.TabStop = False
        Me.fraTxSLConfiguration.Text = "TxSL Configuration"
        '
        'medBarCode
        '
        Me.medBarCode.Location = New System.Drawing.Point(297, 67)
        Me.medBarCode.Name = "medBarCode"
        Me.medBarCode.OcxState = CType(resources.GetObject("medBarCode.OcxState"), System.Windows.Forms.AxHost.State)
        Me.medBarCode.Size = New System.Drawing.Size(33, 10)
        Me.medBarCode.TabIndex = 4
        Me.medBarCode.Visible = False
        '
        'lblBarCode
        '
        Me.lblBarCode.BackColor = System.Drawing.SystemColors.Control
        Me.lblBarCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblBarCode.Enabled = False
        Me.lblBarCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBarCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblBarCode.Location = New System.Drawing.Point(4, 69)
        Me.lblBarCode.Name = "lblBarCode"
        Me.lblBarCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblBarCode.Size = New System.Drawing.Size(123, 19)
        Me.lblBarCode.TabIndex = 25
        Me.lblBarCode.Text = "Bar Code"
        Me.lblBarCode.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblBarCode.Visible = False
        '
        'fraReport
        '
        Me.fraReport.BackColor = System.Drawing.SystemColors.Control
        Me.fraReport.Controls.Add(Me.rtbReport)
        Me.fraReport.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraReport.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraReport.Location = New System.Drawing.Point(384, 136)
        Me.fraReport.Name = "fraReport"
        Me.fraReport.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReport.Size = New System.Drawing.Size(361, 361)
        Me.fraReport.TabIndex = 19
        Me.fraReport.TabStop = False
        Me.fraReport.Text = "Report"
        '
        'rtbReport
        '
        Me.rtbReport.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbReport.Location = New System.Drawing.Point(8, 16)
        Me.rtbReport.MaxLength = 10000
        Me.rtbReport.Name = "rtbReport"
        Me.rtbReport.ReadOnly = True
        Me.rtbReport.Size = New System.Drawing.Size(344, 336)
        Me.rtbReport.TabIndex = 0
        Me.rtbReport.Text = ""
        '
        'fraOperatorMessage
        '
        Me.fraOperatorMessage.BackColor = System.Drawing.SystemColors.Control
        Me.fraOperatorMessage.Controls.Add(Me.rtbViewErrors)
        Me.fraOperatorMessage.Controls.Add(Me.rtbOperatorMessage)
        Me.fraOperatorMessage.Enabled = False
        Me.fraOperatorMessage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraOperatorMessage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraOperatorMessage.Location = New System.Drawing.Point(384, 44)
        Me.fraOperatorMessage.Name = "fraOperatorMessage"
        Me.fraOperatorMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraOperatorMessage.Size = New System.Drawing.Size(361, 89)
        Me.fraOperatorMessage.TabIndex = 17
        Me.fraOperatorMessage.TabStop = False
        Me.fraOperatorMessage.Text = "Operator Messages"
        '
        'rtbViewErrors
        '
        Me.rtbViewErrors.BackColor = System.Drawing.Color.Red
        Me.rtbViewErrors.Font = New System.Drawing.Font("MS Reference Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbViewErrors.Location = New System.Drawing.Point(72, 32)
        Me.rtbViewErrors.Name = "rtbViewErrors"
        Me.rtbViewErrors.ReadOnly = True
        Me.rtbViewErrors.Size = New System.Drawing.Size(176, 40)
        Me.rtbViewErrors.TabIndex = 31
        Me.rtbViewErrors.Text = ""
        Me.rtbViewErrors.Visible = False
        '
        'rtbOperatorMessage
        '
        Me.rtbOperatorMessage.Font = New System.Drawing.Font("MS Reference Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbOperatorMessage.Location = New System.Drawing.Point(8, 16)
        Me.rtbOperatorMessage.Name = "rtbOperatorMessage"
        Me.rtbOperatorMessage.ReadOnly = True
        Me.rtbOperatorMessage.Size = New System.Drawing.Size(344, 64)
        Me.rtbOperatorMessage.TabIndex = 31
        Me.rtbOperatorMessage.Text = ""
        '
        'fraTestplanProgress
        '
        Me.fraTestplanProgress.BackColor = System.Drawing.SystemColors.Control
        Me.fraTestplanProgress.Controls.Add(Me.prbTestplan)
        Me.fraTestplanProgress.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraTestplanProgress.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraTestplanProgress.Location = New System.Drawing.Point(11, 136)
        Me.fraTestplanProgress.Name = "fraTestplanProgress"
        Me.fraTestplanProgress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraTestplanProgress.Size = New System.Drawing.Size(361, 57)
        Me.fraTestplanProgress.TabIndex = 16
        Me.fraTestplanProgress.TabStop = False
        Me.fraTestplanProgress.Text = "Testplan Progress"
        '
        'prbTestplan
        '
        Me.prbTestplan.Location = New System.Drawing.Point(8, 24)
        Me.prbTestplan.Name = "prbTestplan"
        Me.prbTestplan.Size = New System.Drawing.Size(344, 24)
        Me.prbTestplan.TabIndex = 31
        '
        'fraExecution
        '
        Me.fraExecution.BackColor = System.Drawing.SystemColors.Control
        Me.fraExecution.Controls.Add(Me.txtExecutionMode)
        Me.fraExecution.Controls.Add(Me.lblExecutionMode)
        Me.fraExecution.Controls.Add(Me.txtSerialNumber)
        Me.fraExecution.Controls.Add(Me.cmdRun)
        Me.fraExecution.Controls.Add(Me.cmdStop)
        Me.fraExecution.Controls.Add(Me.cmdAbort)
        Me.fraExecution.Controls.Add(Me.lblSerialNumber)
        Me.fraExecution.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraExecution.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraExecution.Location = New System.Drawing.Point(11, 200)
        Me.fraExecution.Name = "fraExecution"
        Me.fraExecution.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraExecution.Size = New System.Drawing.Size(361, 119)
        Me.fraExecution.TabIndex = 15
        Me.fraExecution.TabStop = False
        Me.fraExecution.Text = "Testplan Execution"
        '
        'txtExecutionMode
        '
        Me.txtExecutionMode.AutoSize = True
        Me.txtExecutionMode.Location = New System.Drawing.Point(152, 41)
        Me.txtExecutionMode.Name = "txtExecutionMode"
        Me.txtExecutionMode.Size = New System.Drawing.Size(0, 14)
        Me.txtExecutionMode.TabIndex = 25
        '
        'lblExecutionMode
        '
        Me.lblExecutionMode.AutoSize = True
        Me.lblExecutionMode.Location = New System.Drawing.Point(57, 40)
        Me.lblExecutionMode.Name = "lblExecutionMode"
        Me.lblExecutionMode.Size = New System.Drawing.Size(86, 14)
        Me.lblExecutionMode.TabIndex = 24
        Me.lblExecutionMode.Text = "Execution Mode:"
        '
        'lblSerialNumber
        '
        Me.lblSerialNumber.BackColor = System.Drawing.SystemColors.Control
        Me.lblSerialNumber.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSerialNumber.Enabled = False
        Me.lblSerialNumber.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSerialNumber.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSerialNumber.Location = New System.Drawing.Point(16, 17)
        Me.lblSerialNumber.Name = "lblSerialNumber"
        Me.lblSerialNumber.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSerialNumber.Size = New System.Drawing.Size(126, 19)
        Me.lblSerialNumber.TabIndex = 23
        Me.lblSerialNumber.Text = "Serial Number"
        Me.lblSerialNumber.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fraSystemStatus
        '
        Me.fraSystemStatus.BackColor = System.Drawing.SystemColors.Control
        Me.fraSystemStatus.Controls.Add(Me.rtbSystemStatus)
        Me.fraSystemStatus.Enabled = False
        Me.fraSystemStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraSystemStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraSystemStatus.Location = New System.Drawing.Point(11, 44)
        Me.fraSystemStatus.Name = "fraSystemStatus"
        Me.fraSystemStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraSystemStatus.Size = New System.Drawing.Size(361, 89)
        Me.fraSystemStatus.TabIndex = 13
        Me.fraSystemStatus.TabStop = False
        Me.fraSystemStatus.Text = "System Status"
        '
        'rtbSystemStatus
        '
        Me.rtbSystemStatus.Font = New System.Drawing.Font("Tahoma", 27.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtbSystemStatus.Location = New System.Drawing.Point(8, 16)
        Me.rtbSystemStatus.Name = "rtbSystemStatus"
        Me.rtbSystemStatus.Size = New System.Drawing.Size(344, 64)
        Me.rtbSystemStatus.TabIndex = 0
        Me.rtbSystemStatus.Text = ""
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.SystemColors.Control
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTitle.Location = New System.Drawing.Point(55, 3)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(297, 41)
        Me.lblTitle.TabIndex = 18
        Me.lblTitle.Text = "Agilent TestExec SL"
        '
        'staDescriptions
        '
        Me.staDescriptions.Location = New System.Drawing.Point(0, 529)
        Me.staDescriptions.Name = "staDescriptions"
        Me.staDescriptions.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbpUutName, Me.sbpTestplanName, Me.sbpVariant, Me.sbpCurrentTestName})
        Me.staDescriptions.ShowPanels = True
        Me.staDescriptions.Size = New System.Drawing.Size(758, 23)
        Me.staDescriptions.TabIndex = 31
        '
        'sbpUutName
        '
        Me.sbpUutName.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbpUutName.MinWidth = 150
        Me.sbpUutName.Name = "sbpUutName"
        Me.sbpUutName.Text = "Uut"
        Me.sbpUutName.ToolTipText = "The name of the Uut"
        Me.sbpUutName.Width = 150
        '
        'sbpTestplanName
        '
        Me.sbpTestplanName.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbpTestplanName.MinWidth = 200
        Me.sbpTestplanName.Name = "sbpTestplanName"
        Me.sbpTestplanName.Text = "Testplan"
        Me.sbpTestplanName.ToolTipText = "The name of the loaded testplan."
        Me.sbpTestplanName.Width = 200
        '
        'sbpVariant
        '
        Me.sbpVariant.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.sbpVariant.MinWidth = 150
        Me.sbpVariant.Name = "sbpVariant"
        Me.sbpVariant.Text = "Variant"
        Me.sbpVariant.ToolTipText = "Current Testplan Variant"
        Me.sbpVariant.Width = 150
        '
        'sbpCurrentTestName
        '
        Me.sbpCurrentTestName.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.sbpCurrentTestName.Name = "sbpCurrentTestName"
        Me.sbpCurrentTestName.Text = "Test:"
        Me.sbpCurrentTestName.ToolTipText = "Current Test Name"
        Me.sbpCurrentTestName.Width = 241
        '
        'staStats
        '
        Me.staStats.Location = New System.Drawing.Point(0, 505)
        Me.staStats.Name = "staStats"
        Me.staStats.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbpPassed, Me.sbpFailed, Me.sbpTotal, Me.sbpYield, Me.sbpSince})
        Me.staStats.ShowPanels = True
        Me.staStats.Size = New System.Drawing.Size(758, 24)
        Me.staStats.SizingGrip = False
        Me.staStats.TabIndex = 33
        '
        'sbpPassed
        '
        Me.sbpPassed.MinWidth = 100
        Me.sbpPassed.Name = "sbpPassed"
        Me.sbpPassed.Text = "Passed:"
        Me.sbpPassed.ToolTipText = "The number of Uuts that have passed"
        '
        'sbpFailed
        '
        Me.sbpFailed.MinWidth = 100
        Me.sbpFailed.Name = "sbpFailed"
        Me.sbpFailed.Text = "Failed:"
        Me.sbpFailed.ToolTipText = "The number of Uuts that have failed."
        '
        'sbpTotal
        '
        Me.sbpTotal.MinWidth = 100
        Me.sbpTotal.Name = "sbpTotal"
        Me.sbpTotal.Text = "Total Tested:"
        Me.sbpTotal.ToolTipText = "The total number of Uuts tests."
        '
        'sbpYield
        '
        Me.sbpYield.MinWidth = 100
        Me.sbpYield.Name = "sbpYield"
        Me.sbpYield.Text = "Yield:"
        Me.sbpYield.ToolTipText = "The percent of Uuts tested that have passed"
        '
        'sbpSince
        '
        Me.sbpSince.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.sbpSince.Name = "sbpSince"
        Me.sbpSince.Text = "Since:"
        Me.sbpSince.ToolTipText = "The date and time of the last reset of this data"
        Me.sbpSince.Width = 358
        '
        'TestExecSL1
        '
        Me.TestExecSL1.Enabled = True
        Me.TestExecSL1.Location = New System.Drawing.Point(379, 6)
        Me.TestExecSL1.Name = "TestExecSL1"
        Me.TestExecSL1.OcxState = CType(resources.GetObject("TestExecSL1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.TestExecSL1.Size = New System.Drawing.Size(64, 58)
        Me.TestExecSL1.TabIndex = 24
        Me.TestExecSL1.Visible = False
        '
        'SysBox
        '
        Me.SysBox.BackColor = System.Drawing.Color.Red
        Me.SysBox.Location = New System.Drawing.Point(477, 15)
        Me.SysBox.Name = "SysBox"
        Me.SysBox.Size = New System.Drawing.Size(31, 29)
        Me.SysBox.TabIndex = 34
        Me.SysBox.TabStop = False
        '
        'SysLabel
        '
        Me.SysLabel.AutoSize = True
        Me.SysLabel.Location = New System.Drawing.Point(446, 21)
        Me.SysLabel.Name = "SysLabel"
        Me.SysLabel.Size = New System.Drawing.Size(25, 14)
        Me.SysLabel.TabIndex = 35
        Me.SysLabel.Text = "sys"
        '
        'FixLabel
        '
        Me.FixLabel.AutoSize = True
        Me.FixLabel.Location = New System.Drawing.Point(525, 21)
        Me.FixLabel.Name = "FixLabel"
        Me.FixLabel.Size = New System.Drawing.Size(19, 14)
        Me.FixLabel.TabIndex = 37
        Me.FixLabel.Text = "fix"
        '
        'FixBox
        '
        Me.FixBox.BackColor = System.Drawing.Color.Red
        Me.FixBox.Location = New System.Drawing.Point(556, 15)
        Me.FixBox.Name = "FixBox"
        Me.FixBox.Size = New System.Drawing.Size(31, 29)
        Me.FixBox.TabIndex = 36
        Me.FixBox.TabStop = False
        '
        'DebugCheckBox
        '
        Me.DebugCheckBox.AutoSize = True
        Me.DebugCheckBox.Location = New System.Drawing.Point(611, 21)
        Me.DebugCheckBox.Name = "DebugCheckBox"
        Me.DebugCheckBox.Size = New System.Drawing.Size(56, 18)
        Me.DebugCheckBox.TabIndex = 38
        Me.DebugCheckBox.Text = "debug"
        Me.DebugCheckBox.UseVisualStyleBackColor = True
        '
        'DebugNumericUpDown
        '
        Me.DebugNumericUpDown.Enabled = False
        Me.DebugNumericUpDown.Location = New System.Drawing.Point(664, 19)
        Me.DebugNumericUpDown.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.DebugNumericUpDown.Name = "DebugNumericUpDown"
        Me.DebugNumericUpDown.Size = New System.Drawing.Size(50, 20)
        Me.DebugNumericUpDown.TabIndex = 39
        Me.DebugNumericUpDown.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(758, 552)
        Me.Controls.Add(Me.DebugNumericUpDown)
        Me.Controls.Add(Me.DebugCheckBox)
        Me.Controls.Add(Me.FixLabel)
        Me.Controls.Add(Me.FixBox)
        Me.Controls.Add(Me.SysLabel)
        Me.Controls.Add(Me.SysBox)
        Me.Controls.Add(Me.staStats)
        Me.Controls.Add(Me.staDescriptions)
        Me.Controls.Add(Me.fraTestplanConfiguration)
        Me.Controls.Add(Me.fraTxSLConfiguration)
        Me.Controls.Add(Me.fraReport)
        Me.Controls.Add(Me.fraOperatorMessage)
        Me.Controls.Add(Me.fraTestplanProgress)
        Me.Controls.Add(Me.fraExecution)
        Me.Controls.Add(Me.fraSystemStatus)
        Me.Controls.Add(Me.imglogo)
        Me.Controls.Add(Me.TestExecSL1)
        Me.Controls.Add(Me.lblTitle)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Location = New System.Drawing.Point(19, 27)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Agilent TestExec SL Operator Interface"
        CType(Me.imglogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraTestplanConfiguration.ResumeLayout(False)
        Me.fraTxSLConfiguration.ResumeLayout(False)
        CType(Me.medBarCode, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraReport.ResumeLayout(False)
        Me.fraOperatorMessage.ResumeLayout(False)
        Me.fraTestplanProgress.ResumeLayout(False)
        Me.fraExecution.ResumeLayout(False)
        Me.fraExecution.PerformLayout()
        Me.fraSystemStatus.ResumeLayout(False)
        CType(Me.sbpUutName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpTestplanName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpVariant, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpCurrentTestName, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpPassed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpFailed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpTotal, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpYield, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.sbpSince, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TestExecSL1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SysBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FixBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DebugNumericUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmMain
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmMain
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmMain()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmMain)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    '*****************************************************************************
    'me.frm
    'me.frx
    '
    'The TypicalOpUI example is intended to be an immediately usable
    'operator interface for TxSL based applications.  Source code is
    'provided so that the code can be modified to fit a particular
    'application.
    '
    'Note that the TypicalOpUI example is intended to be robust
    'enough to be used as is, without modification, in production.
    'As such, TypicalOpUI includes several layers of error checking,
    'and extensive management of the state of controls.  The resulting
    'code, while very usable, is probably not the best place to start
    'learning how to do your own TxSL interface.  If you are just getting
    'started learning how to develop your own TxSL interface, we suggest
    'that you start with the "Simple" interface, contained in
    'TestExecSL x.x\Examples\SimpleOpUI.  A thorough examination
    'of the written and online documentation of the TxSL active X control
    'will also be very helpful.
    '
    'The frmMain files are the main code for the TypicalOpUI application.
    'The me.frm file contains the TxSL control, and all of the logic
    'to respond to its many events.  Also included here is the logic to
    'control the states of the controls (for example, to only enable the
    'Run button when appropriate).  The me.frm file is the core
    'piece of code in the application.  Most of the other forms
    'are called as a result of an event from the TxSL control
    'or an event generated by a button or other user interface
    'control on the form.
    '*****************************************************************************

    Private LocalSource As String
    Private mbBarCodeChange As Boolean
    Private txtSerialNumberChange As Boolean

    Private msEntireReportBlock As String 'Used to hold the entire report block,
    'for display at pauses and stops
    Private msReportBlock1 As String 'Used in displaying report blocks
    Private msReportBlock2 As String
    Private msReportBlock3 As String
    Private msReportBlock4 As String

    Public TxSLState As HPTestExecSL.TestplanState 'Used to hold TxSL state.
    'it is used to not allow operations when the state
    'is incorrect

    'TxSLSimpleState is a condensed representation of the TxSL state machine.
    'It specifically recognizes a "paused" state, which is entered after an
    'AfterTestplanPause event is received.  The standard state machine will send
    'an AfterTestplanPauseEvent, but when queried, will report that it is in a TestplanRunning
    'state.  From one behavior, this is correct, as the internal sequencer is still
    'running.  From managing the operator interface, we need to recognize when we are paused.

    'TxSLSimpleState works as follows
    'A form load event will start the machine in a state of NoTestplan
    'An AfterTestplanLoad event will cause a move to the state of NotRun
    'A BeforeTestplanBegin event will cause a move to the state of Running
    'An AfterTestplanPause event will cause a move to the state of Paused
    'An AfterTestplanStop event will cause a move to the state of Stopped.

    Private Enum TxSLSimpleState
        NoTestplan = 100
        NotRun = 101
        running = 102
        paused = 103
        Stopped = 104
    End Enum
    'Here we define a module level, user defined variable as type TxSLSimpleState.
    Private muTxSLSimpleState As TxSLSimpleState

    'A variable to accumulate the logic behind several checks made on
    'the path, the default variant, and the success of selecting the default
    'variant.
    Public gbAfterTPLoadChecksOkay As Boolean

    'A variable to hold the name of which command is being attempted
    'It is set by each of the "buttons".  It will be read by the event handlers for
    'tmrFailure.  It will be reset by tmrFailure, and by any appropriate event handlers
    Private msCommandAttempt As String

    Private mnOldMousePointer As Cursor 'Used to hold the original mouse pointer

    'Constants for use in loading the status bar panels
    'They are useful to keep the code self documenting
    Public Shared ReadOnly iPASSED As Short = 0
    Public Shared ReadOnly iFAILED As Short = 1
    Public Shared ReadOnly iTOTALTESTED As Short = 2
    Public Shared ReadOnly iYIELD As Short = 3
    Public Shared ReadOnly iSINCE As Short = 4
    Public Shared ReadOnly iCURRENTTESTNAME As Short = 0
    Public Shared ReadOnly iTESTPLANNAME As Short = 1
    Public Shared ReadOnly iCURRENTVARIANTNAME As Short = 2
    Public Shared ReadOnly iUUTNAME As Short = 3

    Private Sub chkReportException_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkReportException.CheckStateChanged 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":chkReportException_Click"
        With TestExecSL1.Testplan.Preference
            If chkReportException.CheckState = System.Windows.Forms.CheckState.Checked Then
                .ReportExceptions = True
            Else
                .ReportExceptions = False
            End If
        End With
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub chkReportFailedTests_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkReportFailedTests.CheckStateChanged 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":chkReportFailedTests_Click"
        With TestExecSL1.Testplan.Preference
            If chkReportFailedTests.CheckState = System.Windows.Forms.CheckState.Checked Then
                .ReportFail = True
            Else
                .ReportFail = False
            End If
        End With
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub chkReportPassedTests_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkReportPassedTests.CheckStateChanged 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":chkReportPassedTests_Click"
        With TestExecSL1.Testplan.Preference
            If chkReportPassedTests.CheckState = System.Windows.Forms.CheckState.Checked Then
                .ReportPass = True
            Else
                .ReportPass = False
            End If
        End With
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub chkShowReport_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkShowReport.CheckStateChanged

        If chkShowReport.CheckState = System.Windows.Forms.CheckState.Checked Then
            rtbReport.Visible = True
            fraReport.Visible = True
        Else
            rtbReport.Visible = False
            fraReport.Visible = False
        End If

    End Sub

    Private Sub cmdDebug_Click()

        'Set the UI back to a state in which we can load and run
        AfterSuccessfulLogin()
        ConfigButtonsBeforeRun()
        rtbReport.Text = ""
        rtbReport.SelectionColor = Color.White

    End Sub

    '***********************************************************
    Private Sub cmdLoadTestplan_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLoadTestplan.Click

        frmLoadTestplan.StartSelection()

    End Sub

    '***********************************************************
    Public Sub cmdAbort_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAbort.Click 'TxSLErrorTrapSite
        'Got here for various reasons.  Most likely would have been an operator push
        'of the Abort key on the user interface.  Could also have got here via
        'operator pushes of function keys, or hardware buttons, whose handlers also
        'called this function.
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdAbort_Click"

        'If case is not paused or running, exit the sub
        'Note the Is condition allows us to look at multiple states
        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case Is <> HPTestExecSL.TestplanState.TestplanRunning, HPTestExecSL.TestplanState.TestplanBreakpoint, HPTestExecSL.TestplanState.TestplanPaused, HPTestExecSL.TestplanState.TestplanStepPause
                Beep() 'An invalid state for this command
                Exit Sub
            Case Else
                'continue
        End Select


        'Enable the failure timer.  This timer will allow us to recover
        'if TxSL doesn't respond within the timeout period.
        'Also post an attempting message to the display
        tmrFailure.Enabled = True
        msCommandAttempt = "Abort"
        rtbSystemStatus.SelectionColor = Color.Black
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAttemptingAbort)
        rtbOperatorMessage.Text = ""

        'This code needs to be here, instead of above the preceding code
        'If not, the SystemStatus will always say AttemptingAbort
        cmdAbort.Enabled = False
        cmdStop.Enabled = False
        TestExecSL1.Testplan.Abort()




        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub cmdContinue_Click() 'TxSLErrorTrapSite
        'Note that Continue button is no longer on the user interface.
        'This code is left in to support function keys, and would be used
        'if a Continue was added to the user interface.

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdContinue_Click"
        'Check to make sure that we are in a state that can accept a Continue command.
        'This means that we are NOT "no testplan", or NOT "Running"
        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case Is <> HPTestExecSL.TestplanState.NoTestplan, HPTestExecSL.TestplanState.TestplanRunning
                Beep() 'An invalid state for this command
                Exit Sub
            Case Else
                'continue
        End Select

        ConfigButtonsAfterRun()
        TestExecSL1.Testplan.[Continue]()
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub


    Private Sub cmdPause_Click() 'TxSLErrorTrapSite
        'Note that a Pause key is no longer on the User interface.
        'This code is left in to support function keys, and would be used
        'if a Pause was added to the user interface.
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdPause_Click"

        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case Is <> HPTestExecSL.TestplanState.TestplanRunning
                Beep() 'An invalid state for this command
                Exit Sub
            Case HPTestExecSL.TestplanState.TestplanRunning
                'continue
        End Select
        TestExecSL1.Testplan.Pause()
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub cmdLogin_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLogin.Click

        frmLogin.Show()

    End Sub


    Public Sub cmdRun_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRun.Click 'TxSLErrorTrapSite
        'Got here by someone pressing the run key on the operator interface, or
        'possible by a hardware button has been pressed.  The event sub for the
        'hardware button would have called this sub.
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdRun_Click"

        'Do not proceed if no statement in testplan
        If (TestExecSL1.Testplan.Statements.Count) Then

            'Check to make sure that we are in a state that can accept a run command.
            'This means that we are NOT  in "no testplan", or NOT "Running"

            If (muTxSLSimpleState = TxSLSimpleState.NotRun Or muTxSLSimpleState = TxSLSimpleState.Stopped) Then
                'we are in a "runnable state"
                ConfigButtonsAfterRun()
                'Enable the failure timer.  This timer will allow us to recover
                'if TxSL doesn't respond within the timeout period.
                'Also post an attempting message to the display
                tmrFailure.Enabled = True
                msCommandAttempt = "Run"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAttemptingRun)
                rtbOperatorMessage.Text = ""
                TestExecSL1.Testplan.Run()
            ElseIf (muTxSLSimpleState = TxSLSimpleState.paused) Then
                'we are in a "continuable state"
                ConfigButtonsAfterRun()
                'Enable the failure timer.  This timer will allow us to recover
                'if TxSL doesn't respond within the timeout period.
                'Also post an attempting message to the display
                tmrFailure.Enabled = True
                msCommandAttempt = "Continue"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAttemptingContinue)
                rtbOperatorMessage.Text = ""
                TestExecSL1.Testplan.[Continue]()
            Else
                'we are not in a state for which we can use this button
                Beep()
            End If
        End If

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub


    Private Sub cmdSelectVariant_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSelectVariant.Click

        frmSelectVariant.Show()

    End Sub


    Private Sub cmdStep_Click() 'TxSLErrorTrapSite
        'Note that a step executed as the first command of
        'a testplan will in fact cause a run to be executed, but only
        'one test is run, and the tesplan is then paused, with a return
        'of TESTPLAN_STEP_PAUSE
        'Note that a step button is no longer on the form.  This code has been left in
        'to support function keys, and in the event a Step key is added.
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdStep_Click"

        'Check to make sure that we are in a state that can accept a step command.
        'This means that we are NOT "no testplan", or NOT "Running"
        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case Is <> HPTestExecSL.TestplanState.NoTestplan, HPTestExecSL.TestplanState.TestplanRunning
                Beep() 'An invalid state for this command
                Exit Sub
            Case Else
                'continue
        End Select

        ConfigButtonsAfterRun()
        TestExecSL1.Testplan.Step()
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub


    Public Sub cmdStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStop.Click 'TxSLErrorTrapSite
        'Got here by someone pressing a stop button on the operator interface,
        'or by a hardware button being pressed, and the sub for that hardware button
        'called this sub.
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdStop_Click"

        'Check to make sure that we are in a state that can accept a Stop command.
        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case Is <> HPTestExecSL.TestplanState.TestplanRunning, HPTestExecSL.TestplanState.TestplanBreakpoint, HPTestExecSL.TestplanState.TestplanPaused, HPTestExecSL.TestplanState.TestplanStepPause
                Beep() 'An invalid state for this command
                Exit Sub
            Case Else
                'continue
        End Select

        'Enable the failure timer.  This timer will allow us to recover
        'if TxSL doesn't respond within the timeout period.
        'Also post an attempting message to the display
        tmrFailure.Enabled = True
        msCommandAttempt = "Stop"
        rtbSystemStatus.SelectionColor = Color.Black
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAttemptingStop)
        rtbOperatorMessage.Text = ""

        'This code needed to be here to work correctly.
        'If it was above the preceding block of code, we got events in the wrong order.
        cmdStop.Enabled = False
        TestExecSL1.Testplan.Stop()

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub


    Private Sub cmdTxSLExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTxSLExit.Click 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":cmdTxSLExit_Click"

        frmMain_Closed(Me, New System.EventArgs())
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub frmMain_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        'This sub (when keypreview on the form is also enabled)
        'allows us to intercept keyboard events and handle
        'them appropriate.  It's major use is for the keypad
        'State tracking is also provided in the click events, to
        'reject keypresses when not appropriate.
        Select Case KeyCode
            'These first cases are put in to allow the use of function keys even
            'in the absence of a menu bar.
            Case System.Windows.Forms.Keys.F5
                cmdRun_Click(cmdRun, New System.EventArgs())
            Case System.Windows.Forms.Keys.F7
                If (Shift And VB6.ShiftConstants.ShiftMask) Then
                    cmdAbort_Click(cmdAbort, New System.EventArgs())
                ElseIf (Shift = 0) Then
                    cmdStop_Click(cmdStop, New System.EventArgs())
                Else
                    Beep()
                End If
                'These next cases handle a keypad that has been programmed to generate these
                'obscure button sequences.
            Case System.Windows.Forms.Keys.A
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    cmdRun_Click(cmdRun, New System.EventArgs())
                End If
            Case System.Windows.Forms.Keys.S
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    cmdStop_Click(cmdStop, New System.EventArgs())
                End If
            Case System.Windows.Forms.Keys.D
                'not used for hot key
            Case System.Windows.Forms.Keys.F
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    cmdAbort_Click(cmdAbort, New System.EventArgs())
                End If

                'These cases are supplied for triggering serial bar code readers
                'They are only intended for diagnostics
            Case System.Windows.Forms.Keys.Q
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    frmSerialBarCode.cmdScanOnce_Click(frmSerialBarCode.cmdScanOnce, New System.EventArgs())
                End If
            Case System.Windows.Forms.Keys.W
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    frmSerialBarCode.cmdStartScanContinuous_Click(frmSerialBarCode.cmdStartScanContinuous, New System.EventArgs())
                End If
            Case System.Windows.Forms.Keys.E
                If (Shift = (VB6.ShiftConstants.ShiftMask + VB6.ShiftConstants.CtrlMask + VB6.ShiftConstants.AltMask)) Then
                    frmSerialBarCode.cmdStopScanContinous_Click(frmSerialBarCode.cmdStopScanContinous, New System.EventArgs())
                End If
            Case Else
                'do nothing
        End Select

    End Sub


    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":Form_Load"

        'Show the splash screen while loading
        frmSplash.Show()
        frmSplash.Refresh()


        'check to see if we have any testplans to load
        'if not, we show an error, and exit
        Dim RegisteredTestplanInstance As Object
        RegisteredTestplanInstance = Me.TestExecSL1.RegisteredTestplans.Item(1)
        If RegisteredTestplanInstance Is Nothing Then
            Err.Raise(frmErrorDialog.TxSLOpuiError.EmptyTestplanRegistry, LocalSource, "There are no testplans registered to run in the Preferences.upf file.  The operator interface will exit.")
        End If

        'Set the configuration variables
        modConfiguration.Configure()

        'Initialize the language array to all languages available
        modLocalization.InitializeLangArray()
        modAppSpecific.UpdateControlCaptions()

        Show() 'shows the main form

        'this code simply calls a TxSL property.   The code will block until it returns.
        'the net result is once the code does return, I can unload the splash screen.
        Dim WaitForCompletionFlag As Boolean

        WaitForCompletionFlag = TestExecSL1.Security.IsAUserLoggedIn

        'Manage configuration
        If gbTxSLSharedIO = True Then
            gfrmTxSLSharedIO = New frmTxSLSharedIO
        End If
        If gbSerialBoxBarCode = True Then
            gfrmSerialBoxBarCode = New frmSerialBarCode
        End If
        If gbStripPrinterAvailable = True Then
            gfrmStripPrinter = New frmStripPrinter
        End If

        'Call the login screen
        frmLogin.Show() 'A successful login will change the ui appropriately

        'set the simple state of TxSL
        muTxSLSimpleState = TxSLSimpleState.NoTestplan

        'Unload the splash screen
        frmSplash.Close()

        'set the alignment of the status box.  This must be set here, as it is not
        'available at design time
        rtbSystemStatus.SelectionAlignment = HorizontalAlignment.Center
        rtbViewErrors.SelectionAlignment = HorizontalAlignment.Center

        Exit Sub

LocalErrorHandler:
        'If we encountered errors on the load, we post a dialog
        'and exit.  We can't go on
        MsgBox(Err.Description)
        frmMain_Closed(Me, New System.EventArgs)


    End Sub

    Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":Form_Unload"

        Dim tpa As HPTestCoreRuntime.Testplan = TestExecSL1.Testplan
        If tpa Is Nothing Then
            TxSLState = HPTestExecSL.TestplanState.NoTestplan
        Else
            TxSLState = TestExecSL1.Testplan.State
        End If

        Dim MsgResult As Short
        Dim Msg As String
        Select Case TxSLState

            Case HPTestExecSL.TestplanState.TestplanRunning, HPTestExecSL.TestplanState.TestplanPaused, HPTestExecSL.TestplanState.TestplanStepPause, HPTestExecSL.TestplanState.TestplanBreakpoint

                Msg = "The test system is either Running, Paused, waiting at a breakpoint or waiting for a step.  " & vbCrLf & vbCrLf & "Pressing OK will result in the operator interface stopping, but the test system may continue to run, or it may halt with power still applied.  " & "Either condition may be hazardous.  " & vbCrLf & "Press OK if this behavior is acceptable.  " & vbCrLf & vbCrLf & "Pressing Cancel will result in control returning to the operator interface, where you can attempt a safe shutdown.  " & "However, it is likely that some data has been lost.  " & "TxSL will attempt to synchronize the state of testplan execution and the state of the operator interface.  " & vbCrLf & vbCrLf & "Press Cancel if this behavior is acceptable"

                MsgResult = MsgBox(Msg, MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
                If MsgResult = MsgBoxResult.Cancel Then
                    tmrFailure.Enabled = False
                    Synch()
                    Exit Sub
                Else
                    TestExecSL1.Testplan.Abort()
                End If

                'If it is vbok, we will simply fall out of this case statement, and execute
                'the unload.

            Case Else
                'we can safely exit from this state, so we will do nothing, and keep on going

        End Select

        rtbSystemStatus.SelectionColor = Color.Black
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAttemptingExit)
        rtbOperatorMessage.Text = ""

        'TestExecSL1.UnloadTestplan 'we don't do this, , as it may running
        'and unloading a running testplan will result in an error
        'unload all the forms.  Note that this is done from high to low,
        'to avoid errors in unloading
        If Not gfrmTxSLSharedIO Is Nothing Then gfrmTxSLSharedIO.Close()
        If Not gfrmSerialBoxBarCode Is Nothing Then gfrmSerialBoxBarCode.Close()
        If Not gfrmStripPrinter Is Nothing Then gfrmStripPrinter.Close()

        Application.Exit()

        'if we have set any local references to testplans, they need to be set to nothing here
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub medBarCode_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles medBarCode.Change

        mbBarCodeChange = True

    End Sub


    Private Sub medBarCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles medBarCode.Leave

        'This will be called when the control loses focus.
        If mbBarCodeChange Then
            If medBarCode.CtlText <> "" Then
                HandleBarCode((medBarCode.CtlText))
            End If
            mbBarCodeChange = False
        End If

    End Sub

    Sub medBarCode_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSMask.MaskEdBoxEvents_KeyPressEvent) Handles medBarCode.KeyPressEvent
        'Interprets the Enter key as a tab.  This allows default barcode readers
        'to be used,and directly read into the application, with no changes
        If eventArgs.KeyAscii = System.Windows.Forms.Keys.Return Then
            System.Windows.Forms.SendKeys.Send("{tab}")
            eventArgs.KeyAscii = 0
        End If

    End Sub

    Private Sub medBarCode_ValidationError(ByVal eventSender As System.Object, ByVal eventArgs As AxMSMask.MaskEdBoxEvents_ValidationErrorEvent) Handles medBarCode.ValidationError

        MsgBox(LangLookup(modLocalization.txslLangIndex.gnUnMatchedBarCode))

    End Sub

    Public Sub HandleBarCode(ByRef BarCode As String) 'TxSLErrorTrapSite
        LocalSource = ":" & Me.Name & ":HandleBarCode" 'Defines the context, for use in error handling
        On Error GoTo LocalErrorHandler 'Turns on error handling

        'Customize:
        'A general purpose bar code handler routine.  Note that it needs to be customized for specific
        'applications.

        Dim BarCodeLength As Short
        BarCodeLength = 16 'This value should be changed to whatever is
        'of the input bar code

        'Check for a max length of 64.  This is a limit of the masked edit
        'text box
        If Len(BarCode) > 64 Then
            Err.Raise(frmErrorDialog.TxSLOpuiError.TooLongBarcode, LocalSource, "The maximum length of a bar code is 64 characters.  The bar code you entered had " & CStr(Len(BarCode)) & " characters.")
        End If

        'In case this was called by the serial bar code reader, we will update the
        'bar code control, but set the change variable to false.
        medBarCode.CtlText = BarCode
        mbBarCodeChange = False

        'Disable the input box, to reject any new inputs while processing this one
        medBarCode.Enabled = False
        Dim UUT_Type As String
        Dim UUT_SerialNumber As String
        Dim TestplanPath As String
        Dim DefaultVariant As String
        Dim UUT_Description As String

        'Check to make sure that we are in a state that can accept a bar code
        'Since our assumption is that we will use the bar code to load a testplan,
        'We will only accept a bar code if we are NOT running or paused.
        TxSLState = TestExecSL1.Testplan.State
        Select Case TxSLState
            Case HPTestExecSL.TestplanState.TestplanRunning, HPTestExecSL.TestplanState.TestplanBreakpoint, HPTestExecSL.TestplanState.TestplanPaused, HPTestExecSL.TestplanState.TestplanStepPause
                Err.Raise(frmErrorDialog.TxSLOpuiError.ImproperState, LocalSource, txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnImproperState), CStr(TxSLState)))
            Case Else
                'continue
        End Select

        'You may wish to check for the proper bar code format here.  Such
        'checks might include the length of the string, the proper terminator etc.
        'Here the test is for proper length
        'We raise an error if the length is not correct.
        If Len(BarCode) <> BarCodeLength Then
            Err.Raise(frmErrorDialog.TxSLOpuiError.ImproperBarCodeLength, LocalSource, txslFormatString2(LangLookup(modLocalization.txslLangIndex.gnImproperBarCodeLength), BarCode, CStr(BarCodeLength)))
        End If


        'MODIFIFY HERE to accomodate different packing of information into the
        'bar code string.
        'Note that this parsing code will need to be modified to reflect
        'customer specific requirements.
        'Also note that the following approach was based on the bar code information being position dependent,
        'and not delimited.  Delimited approaches are possible, but will take entirely different
        'parsing approaches.

        UUT_Type = Mid(BarCode, 1, 7) 'Starting at position 1, take the first 7 characters
        UUT_SerialNumber = Mid(BarCode, 8, 9) 'Starting at position 8, take the next 9 characters

        'Use the UUT type to look up the TestplanIndex from the RegisteredUUTs collection
        'Then use that TestplanIndex to lookup the testplan path from the RegisteredTestplans collection
        'That path will then be used to load the testplan.
        Dim TestplanIndex As Short
        TestplanIndex = Me.TestExecSL1.RegisteredUUTs.Item(UUT_Type).TestplanIndex
        TestplanPath = Me.TestExecSL1.RegisteredTestplans.Item(TestplanIndex).Path
        DefaultVariant = Me.TestExecSL1.RegisteredTestplans.Item(TestplanIndex).DefaultVariant

        'Load the testplan.  If it is already loaded, it will return immediately
        If frmLoadTestplan.LoadTestplan(TestplanPath, DefaultVariant) Then
            'Note that the AfterTestplanLoad event will update the screens as appropriate.
            'Clear the medBarCode box, so that it can be used again.
            medBarCode.Enabled = True
            medBarCode.CtlText = ""

            'Update the txtbox that contains the serial number, and the system symbol table.
            txtSerialNumber.Text = UUT_SerialNumber
            UpdateTxSLSerialNumberFields(UUT_SerialNumber)

            'Update the description
            staDescriptions.Panels(iUUTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUUT1Param), (Me.TestExecSL1.RegisteredUUTs.Item(UUT_Type).Description))

            'If gbRunAutomaticallyAfterBarCode is enabled, go ahead and run the testplan
            If gbRunAutomaticallyAfterBarCode = True Then
                cmdRun_Click(cmdRun, New System.EventArgs)
            End If
        Else
            'The load didn't succeed, and warnings were done in the LoadTestplan sub
            'Simply exit, with no changes
        End If
        Exit Sub
        'Error handling

LocalErrorHandler:
        Select Case Err.Number
            Case 91
                'Error 91 is the most likely error if the UUT_Type value can not be found in the
                'UUT Registry, as defined in the tstexcls.ini file

                frmErrorDialog.ErrorHandler((frmErrorDialog.TxSLOpuiError.UnRecognizedBarCode), Err.Source & LocalSource, txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUUTRegistryError), UUT_Type))
                ConfigControlsAfterLoadFailed()

                'move the focus back to the error form
                'NewErrorDialog.cmdOK.SetFocus

            Case frmErrorDialog.TxSLOpuiError.ImproperBarCodeLength
                frmErrorDialog.ErrorHandler(Err.Number, Err.Source, Err.Description)
                ConfigControlsAfterLoadFailed()

                'move the focus back to the error form
                'NewErrorDialog.cmdOK.SetFocus

            Case frmErrorDialog.TxSLOpuiError.ImproperState
                'Could have actually raised an error, but this was viewed as too harsh
                'The following code would have raised the error
                'frmErrorDialog.ErrorHandler Err.Number, Err.Source, Err.Description
                'ConfigControlsAfterLoadFailed

                'Instead, just reenable the bar code field, (so that it catches any others inputs
                'beep and exit
                medBarCode.Enabled = True
                Beep()
                Exit Sub

            Case Else
                frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
        End Select

    End Sub

    Private Sub TestExecSL1_AdviseUpdate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TestExecSL1.AdviseUpdate
        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("Advise Update at : " & TimeOfDay)

        UpdateOperatorInterface()
    End Sub

    Private Sub TestExecSL1_AfterTestDone(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_AfterTestDoneEvent) Handles TestExecSL1.AfterTestDone
        'This event would be used to create custom reports and data logging schemes.
        'Implementation is up to the programmer.
        'Remember that events at the test level incur cross process communication for
        'each event, and may result in prohibitive overhead for time critical applications.
        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AfterTestDone:  TestName = " & eventArgs.testName)

        UpdateOperatorInterface()

    End Sub

    Private Sub TestExecSL1_BeforeTestBegin(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_BeforeTestBeginEvent) Handles TestExecSL1.BeforeTestBegin
        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("BeforeTestBegin:  TestName = " & eventArgs.testName)

    End Sub

    Private Sub TestExecSL1_UserDefinedMessage(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_UserDefinedMessageEvent) Handles TestExecSL1.UserDefinedMessage
        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("UserDefinedMessage:  Message ID = " & eventArgs.id & "  Message =: " & eventArgs.message)

        'Customize:
        'A sub to handle messages that were sent to the operator interface.
        'This routine first checks for HP predefined messages, and reacts appropriately
        'If no match is found, then we simply go on.
        If frmPreDefinedTxSLUserMessages.UserDefinedMessage(eventArgs.id, eventArgs.message) Then
            'If the above call returns true, the message was handled by the
            'frmPreDefinedUserMessage form

            'elseif
            'add custom message handlers here
        Else
            'An unknown message was received. We will do nothing and go on.
            'To diagnose this, turn on set gbTxSLDebugEvents = True in the modConfiguration.bas
            'file and watch the output of the UserMessageEvent.
            'If it is a message that you wish to handle, then add it to the case statement above
            'and add code to handle it.

        End If

    End Sub


    Private Sub TestExecSL1_AfterTestplanLoad(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_AfterTestplanLoadEvent) Handles TestExecSL1.AfterTestplanLoad 'TxSLErrorTrapSite
        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AfterTestPlanLoad:  Path = " & eventArgs.path)
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":TestExecSL1_AfterTestplanLoad"

        'Update the local tracking of TxSL state
        muTxSLSimpleState = TxSLSimpleState.NotRun

        'Now that we have a testplan, we can populate the variants box
        frmSelectVariant.PopulateVariantsBox()

        'Just about to enter a three if statements, that check to see if we really want to go on.
        If frmLoadTestplan.PathIsValid(eventArgs.path) = False Then
            'The path was not valid, and we will exit
            ConfigControlsAfterLoadFailed()
            gbAfterTPLoadChecksOkay = False
            Exit Sub
        End If


        Dim DefaultVariant As String
        DefaultVariant = frmLoadTestplan.RetrieveDefaultVariant(eventArgs.path)
        If frmSelectVariant.VariantIsValid(DefaultVariant) = False Then
            'The variant indexed by this path is not valid, and we will exit
            ConfigControlsAfterLoadFailed()
            gbAfterTPLoadChecksOkay = False
            Exit Sub
        End If

        If frmSelectVariant.SelectVariant(DefaultVariant) = False Then
            'The attempt to select the variant failed, and we will exit
            ConfigControlsAfterLoadFailed()
            gbAfterTPLoadChecksOkay = False
            Exit Sub
        End If




        'Success, by getting to here we have
        '1) Successfully loaded a testplan (proved by the AfterTestplanLoad event)
        '2) Checked that the testplan that we loaded was in the list of
        'registered testplans
        '3) Checked that the default variant name is valid, by checking in the
        'Variants collections
        '4) and (finally) successfully loaded that variant

        gbAfterTPLoadChecksOkay = True
        frmLoadTestplan.UpdateCaptions((eventArgs.path))

        'We can now go on to set all the controls up, and proceed to running.

        'update the TestplanConfiguration  options
        fraTestplanConfiguration.Enabled = True
        chkShowReport.Enabled = True
        With TestExecSL1.Testplan.Preference
            If .ReportExceptions Then
                chkReportException.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkReportException.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            chkReportException.Enabled = True
            If .ReportFail Then
                chkReportFailedTests.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkReportFailedTests.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            chkReportFailedTests.Enabled = True
            If .ReportPass Then
                chkReportPassedTests.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkReportPassedTests.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            chkReportPassedTests.Enabled = True
            .FailCountLimit = DebugNumericUpDown.Value
        End With

        'update the statistics status bar
        With staStats
            .Panels(iPASSED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnPassed1Param), CStr(TestExecSL1.Testplan.History.PassedUUTCount))
            .Panels(iFAILED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnFailed1Param), CStr(TestExecSL1.Testplan.History.FailedUUTCount))
            .Panels(iTOTALTESTED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTested1param), CStr(TestExecSL1.Testplan.History.TestedUUTCount))
            .Panels(iYIELD).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnYield1param), VB6.Format(YieldCalculation, "Percent"))
            .Panels(iSINCE).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnSince1param), VB6.Format(Now, "General Date"))
        End With

        With staDescriptions
            .Panels(iCURRENTTESTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), "")
            .Panels(iTESTPLANNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTestplan1Param), (frmLoadTestplan.txtCurrentTestplanName.Text))
            .Panels(iCURRENTVARIANTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnVariant1Param), (frmSelectVariant.rtbCurrentVariant.Text))
            .Panels(iUUTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUUT1Param), (TestExecSL1.SymbolTables.Item("System").Symbols.Item("ModuleType").Value))
        End With

        'clear the report box, so that no stale data is shown
        Me.rtbReport.Text = ""
        Me.rtbReport.BackColor = System.Drawing.Color.White

        'Update testplan execution mode
        If TestExecSL1.Testplan.Threading Then
            Me.txtExecutionMode.Text = "Threading"
        Else
            Me.txtExecutionMode.Text = "Sequential"
        End If

        'Set the buttons in a runnable state
        ConfigButtonsBeforeRun()

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Private Sub ConfigButtonsAfterRun()
        'A helper function that disables parts of the UI, based on
        'the assumption that a command to run a testplan will be successful.
        'Note that UI elements are disabled, but none are enabled.
        'The before testplan begin sub will enable the Stop and Abort keys

        cmdRun.Enabled = False
        'cmdStep.Enabled = False
        'cmdContinue.Enabled = False
        'disable some of the TestplanConfiguration options.  Note that ShowReport is in
        'fact a property of this program (and not of the TxSL control).  It is left
        'enabled
        chkReportPassedTests.Enabled = False
        chkReportFailedTests.Enabled = False
        chkReportException.Enabled = False

        'disable the ability to modify the serial number while running
        txtSerialNumber.Enabled = False
        txtSerialNumber.BackColor = System.Drawing.SystemColors.InactiveBorder
        lblSerialNumber.Enabled = False

        'disable the TxSL configuration area
        cmdLogin.Enabled = False
        cmdLoadTestplan.Enabled = False
        cmdSelectVariant.Enabled = False

        'disable any more inputs into the bar code
        medBarCode.Enabled = False

    End Sub


    Public Sub ConfigButtonsBeforeRun()

        'a helper routine that sets up buttons in a runnable state
        'it is called whenever we wish to set the buttons to run
        'For example, it would be called from a successful exit result of LoadTestplan, from
        'AfterTestplanPause and AfterTestplanStop events.

        fraExecution.Enabled = True
        'cmdStep.Enabled = True
        'cmdContinue.Enabled = False
        'cmdPause.Enabled = False
        cmdStop.Enabled = False
        cmdAbort.Enabled = False
        cmdRun.Text = LangLookup(modLocalization.txslLangIndex.gnRun1)
        cmdRun.Enabled = True
        cmdRun.Visible = True

        If gbWedgeKeyboardBarCode Then
            medBarCode.Enabled = True
            medBarCode.Focus()
            medBarCode.CtlText = ""
        Else
            cmdRun.Focus()
        End If

        'update the TestplanExecution options
        lblSerialNumber.Enabled = True
        txtSerialNumber.Enabled = True
        txtSerialNumber.BackColor = System.Drawing.Color.White
        InitSerialInput()

        'change the active test name to nothing
        staDescriptions.Panels(iCURRENTTESTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), "")

        'update the progress bar to zero
        prbTestplan.Value = 0

        'configure the operator message box
        Me.rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnReadyToTest)

        _readyToRunEvent.Set()
    End Sub

    Private Sub InitSerialInput()
        ' Set focus to the input box and clear it
        txtSerialNumber.Focus()
        txtSerialNumber.Text = ""
    End Sub

    Private Sub TestExecSL1_AdviseClearReport(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TestExecSL1.AdviseClearReport

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AdviseClearReport")

        'configure the report box
        rtbReport.Text = ""
        rtbReport.BackColor = System.Drawing.Color.White

        msEntireReportBlock = ""
        msReportBlock1 = ""
        msReportBlock2 = ""
        msReportBlock3 = ""
        msReportBlock4 = ""

    End Sub

    Private Sub TestExecSL1_AfterTestplanUnload(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TestExecSL1.AfterTestplanUnload 'TxSLErrorTrapSite

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AfterTestplanUnload")
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":TestExecSL1_AfterTestplanUnload"

        'Manage Configurations
        If gbTxSLSharedIO = True Then frmTxSLSharedIO.AfterTestplanUnload()

        'No changes to serial bar code
        'Manage the forms and menu's associates with the large demo

        rtbSystemStatus.SelectionColor = Color.Cyan
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnEmpty)
        prbTestplan.Value = 0

        'reset the TestplanExecution Area
        cmdStop.Enabled = False
        cmdRun.Enabled = False
        cmdAbort.Enabled = False
        cmdLoadTestplan.Enabled = True
        'cmdLoadTestplan.SetFocus   done below
        lblSerialNumber.Enabled = False
        txtSerialNumber.Enabled = False
        txtSerialNumber.Text = ""

        'set the TestplanOptions to Unchecked, and disable them
        fraTestplanConfiguration.Enabled = False
        chkReportException.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkReportException.Enabled = False
        chkReportFailedTests.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkReportFailedTests.Enabled = False
        chkReportPassedTests.CheckState = System.Windows.Forms.CheckState.Unchecked
        chkReportPassedTests.Enabled = False
        cmdSelectVariant.Enabled = False

        With staStats
            .Panels(iPASSED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnPassed1Param), "")
            .Panels(iFAILED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnFailed1Param), "")
            .Panels(iTOTALTESTED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTested1param), "")
            .Panels(iYIELD).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnYield1param), "")
            .Panels(iSINCE).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnSince1param), "")
        End With

        With staDescriptions
            .Panels(iCURRENTTESTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), "")
            .Panels(iTESTPLANNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTestplan1Param), "")
            .Panels(iCURRENTVARIANTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnVariant1Param), "")
            .Panels(iUUTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUUT1Param), (TestExecSL1.SymbolTables.Item("System").Symbols.Item("ModuleType").Value))
        End With
        rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnPleaseLoadTestplan)
        rtbReport.Text = ""
        rtbReport.BackColor = System.Drawing.Color.White

        'Reenable the barcode entry box
        medBarCode.Enabled = True

        'set the focus appropriately
        If gbWedgeKeyboardBarCode Then
            medBarCode.Focus()
        Else
            cmdLoadTestplan.Focus()
        End If

        'Clear testplan execution mode
        Me.txtExecutionMode.Text = ""

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    '***************************************************************************
    Private Sub TestExecSL1_BeforeTestplanLoad(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_BeforeTestplanLoadEvent) Handles TestExecSL1.BeforeTestplanLoad

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("BeforeTestplanLoad Path = " & eventArgs.path)
        'if you have set any local references to the testplan, they should be set to nothing here
        rtbSystemStatus.SelectionColor = Color.Black
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnLoadingTestplan)
        rtbOperatorMessage.Text = ""
        frmLoadTestplan.BeforeTestplanLoad()

    End Sub

    '***************************************************************************
    Private Sub TestExecSL1_ReportMessage(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_ReportMessageEvent) Handles TestExecSL1.ReportMessage

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("ReportMessage:" & vbCrLf & eventArgs.message)
        'simply append the incoming reportblock onto the end of the existing
        'report contents.


        'Experiments (in VB6, not vb.net) have shown that keeping a running accumulation of report events in
        'rtbReport.txt can slow the testplan, especially when many report messages are
        'arriving.
        'This is true even if using an approach where the .selstart position is
        'set to a very large number and insertions made there.
        'A slighly faster alternative is to only show the last 4 messages (when running),
        'accumulate the entire report block in a string, and only show the
        'entire string (when stopped).
        'The following code only shows the last four messages, but still accumulates
        'all the report messages.

        'It is essentially a four string shift register
        'old way
        'msReportBlock1 = msReportBlock2
        'msReportBlock2 = msReportBlock3
        'msReportBlock3 = msReportBlock4
        'msReportBlock4 = eventArgs.message
        'rtbReport.Text = msReportBlock1 & msReportBlock2 & msReportBlock3 & msReportBlock4

        'This code accumulates the reports into the report block variable
        msEntireReportBlock = msEntireReportBlock & eventArgs.message

        'this code will autoscroll the report window
        Dim ctr As Control = Me.ActiveControl
        With Me.rtbReport
            .Focus()
            .AppendText(eventArgs.message)
            .ScrollToCaret()
        End With
        ctr.Focus()
        If (rtbReport.Lines.Length > rtbReport.MaxLength) Then
            rtbReport.Select(0, rtbReport.GetFirstCharIndexFromLine(rtbReport.Lines.Length - rtbReport.MaxLength))
            rtbReport.SelectedText = vbNullChar
            msEntireReportBlock = rtbReport.Text
        End If

    End Sub


    Private Sub TestExecSL1_AfterTestplanStop(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_AfterTestplanStopEvent) Handles TestExecSL1.AfterTestplanStop 'TxSLErrorTrapSite

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AfterTestplanStop:  Reason = " & eventArgs.reason)
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":TestExecSL1_AfterTestplanStop"

        'Update the local tracking of TxSL state
        muTxSLSimpleState = TxSLSimpleState.Stopped

        'Manage different configurations
        If gbTxSLSharedIO = True Then frmTxSLSharedIO.AfterTestplanStop(eventArgs.reason)

        'Disable the failure timer, as we have responded to the Stop or Abort buttons
        tmrFailure.Enabled = False

        'Enable the bar code areas
        medBarCode.Enabled = True
        lblBarCode.Enabled = True
        medBarCode.BackColor = System.Drawing.Color.White

        'enable login, load of a new testplan, and exit.
        ConfigButtonsAfterStop()

        'enable the ability to change the TestplanConfiguration options
        chkReportPassedTests.Enabled = True
        chkReportFailedTests.Enabled = True
        chkReportException.Enabled = True

        'enable the serial number area
        txtSerialNumber.Enabled = True
        txtSerialNumber.BackColor = System.Drawing.Color.White
        lblSerialNumber.Enabled = True

        'Set the progress bar to max, then disable, and set to zero.
        prbTestplan.Value = prbTestplan.Maximum
        fraTestplanProgress.Enabled = False
        prbTestplan.Value = 0
        prbTestplan.Enabled = False

        'Update the status bar
        staStats.Panels(iPASSED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnPassed), CStr(TestExecSL1.Testplan.History.PassedUUTCount))
        staStats.Panels(iPASSED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnPassed1Param), CStr(TestExecSL1.Testplan.History.PassedUUTCount))
        staStats.Panels(iFAILED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnFailed1Param), CStr(TestExecSL1.Testplan.History.FailedUUTCount))
        staStats.Panels(iTOTALTESTED).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTested1param), CStr(TestExecSL1.Testplan.History.TestedUUTCount))
        staStats.Panels(iYIELD).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnYield1param), VB6.Format(YieldCalculation, "Percent"))

        'Leave the testname panel displayed, so that we can see the last test executed
        'Dim temp As String = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), TestExecSL1.Testplan.CurrentTestExecuting)
        staDescriptions.Panels(iCURRENTTESTNAME).Text = ""

        'update the system status display, and the operator message display
        Select Case eventArgs.reason
            Case HPTestExecSL.TestplanState.TestplanPassed
                rtbSystemStatus.SelectionColor = Color.Lime
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnPassed)
                rtbReport.BackColor = System.Drawing.Color.Lime
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnPassedOpMsg)
                'Manage pass fail indicators
                If gbTxSLSharedIO = True Then frmTxSLSharedIO.AfterTestplanStop(eventArgs.reason)
            Case HPTestExecSL.TestplanState.TestplanFailed
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnFailed)
                rtbReport.BackColor = System.Drawing.Color.Red
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnFailedOpMsg)
            Case HPTestExecSL.TestplanState.TestplanStopped
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnStopped)
                rtbReport.BackColor = System.Drawing.Color.Red
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnStoppedOpMsg)
            Case HPTestExecSL.TestplanState.TestplanAbort
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnAborted)
                rtbReport.BackColor = System.Drawing.Color.Red
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnAbortedOpMsg)
            Case HPTestExecSL.TestplanState.TestplanException
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnException)
                rtbReport.BackColor = System.Drawing.Color.Red
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnExceptionOpMsg)
            Case HPTestExecSL.TestplanState.TestplanUnhandledException
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnUnRecoverableException)
                rtbReport.BackColor = System.Drawing.Color.Red
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnUnrecoverableExceptionOpMsg)
            Case 99
                'This sub was called after a dialog.  We are doing our best to synch up
                'See the synch routine
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = "Synch Up"
                rtbOperatorMessage.Text = ""
            Case Else
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnError)
                rtbReport.BackColor = System.Drawing.Color.White
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnErrorExitOpMsg)
                MsgBox(LangLookup(modLocalization.txslLangIndex.gnErrorExitOpMsg), MsgBoxStyle.OkOnly, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
        End Select

        'Update the report block with the entire contents
        rtbReport.Text = msEntireReportBlock

        ' Calc the name of the log file
        Dim dirName As String = "C:\TestExecLogs\" & Regex.Match(TestExecSL1.Testplan.Path, "[^\\]+(?=\.tpa)").Value
        Dim fileName As String = Date.Now.ToString("yyyyMMdd-HHmmss_") & txtSerialNumber.Text & ".txt"
        Dim filePath As String = dirName & "\" & fileName

        ' Save to file
        My.Computer.FileSystem.CreateDirectory(dirName)
        Using sw As New StreamWriter(File.Open(filePath, FileMode.OpenOrCreate))
            sw.WriteLine(msEntireReportBlock)
        End Using

        If gbTxSLTestplanTimeEnabled Then
            gdtTestplanStopTime = TimeOfDay
            rtbReport.Text = rtbReport.Text & "Total test time: " & VB6.Format(System.DateTime.FromOADate(gdtTestplanStopTime.ToOADate - gdtTestplanStartTime.ToOADate), "h:mm:ss")
        End If
        rtbReport.SelectionStart = MaxSelLength

        'Print to the strip printer, based on availability of the printer, and what flags are set
        If gbStripPrinterAvailable = True Then
            If gbAutoPrintPass = True And eventArgs.reason = HPTestExecSL.TestplanState.TestplanPassed Then
                frmStripPrinter.AfterTestplanStop((msEntireReportBlock))
            End If
            If gbAutoPrintNonPass = True And eventArgs.reason <> HPTestExecSL.TestplanState.TestplanPassed Then
                frmStripPrinter.AfterTestplanStop((msEntireReportBlock))
            End If
        End If

        'finally, set up the focus appropriately on the user interface
        Me.Cursor = mnOldMousePointer
        ConfigButtonsBeforeRun()

        FixBox.BackColor = IIf(_tcpClient.Report(), Color.Red, Color.Yellow)
        InitSerialInput()
        _testDoneEvent.Set()
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    '***************************************************************************
    Private Sub TestExecSL1_AfterTestplanPause(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_AfterTestplanPauseEvent) Handles TestExecSL1.AfterTestplanPause 'TxSLErrorTrapSite

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("AfterTestplanPause:  Reason = " & eventArgs.reason & "  TestName =+ TestName")
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":TestExecSL1_AfterTestplanPause"

        'Update the local tracking of TxSL state
        muTxSLSimpleState = TxSLSimpleState.paused

        staDescriptions.Panels(iCURRENTTESTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), eventArgs.testName)
        'update the system status display
        Select Case eventArgs.reason
            Case HPTestExecSL.TestplanState.TestplanPaused
                rtbSystemStatus.SelectionColor = Color.Blue
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnPaused)
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnPausedOpMsg)
                rtbReport.BackColor = System.Drawing.Color.White
            Case HPTestExecSL.TestplanState.TestplanBreakpoint
                rtbSystemStatus.SelectionColor = Color.Blue
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnBreakpoint)
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnBreakPointOpMsg)
                rtbReport.BackColor = System.Drawing.Color.White
            Case HPTestExecSL.TestplanState.TestplanStepPause
                rtbSystemStatus.SelectionColor = Color.Blue
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnStepPause)
                rtbReport.BackColor = System.Drawing.Color.White
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnStepPauseOpMsg)
            Case 99
                'This sub was called after a dialog.  We are doing our best to synch up
                'See the synch routine
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = "Synch Up"
                rtbOperatorMessage.Text = ""
            Case Else
                rtbSystemStatus.SelectionColor = Color.Red
                rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnPauseError)
                rtbReport.BackColor = System.Drawing.Color.White
                rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnPauseErrorOpMsg)
                MsgBox(LangLookup(modLocalization.txslLangIndex.gnPauseErrorOpMsg), MsgBoxStyle.OkOnly, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
        End Select

        'disable the progress bar
        prbTestplan.Enabled = False

        'set up the buttons
        cmdRun.Text = LangLookup(modLocalization.txslLangIndex.gnContinue1)
        cmdRun.Enabled = True
        'cmdStep.Enabled = True
        'cmdPause.Enabled = False
        cmdStop.Enabled = True
        cmdAbort.Enabled = True

        'No serial bar code changes
        cmdRun.Focus() 'Here the run button is also being used as continue
        'Update the report box, with the entire contents of the report stream
        rtbReport.Text = msEntireReportBlock
        rtbReport.SelectionStart = MaxSelLength

        Me.Cursor = mnOldMousePointer
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    '***************************************************************************
    Private Sub TestExecSL1_BeforeTestplanBegin(ByVal eventSender As System.Object, ByVal eventArgs As AxHPTestExecSL.DTestExecSLEvents_BeforeTestplanBeginEvent) Handles TestExecSL1.BeforeTestplanBegin 'TxSLErrorTrapSite
        'a long sub that configures the human interface
        'into the correct configuration following
        'the start of a testplan.
        'Dim nMaxTests As Short
        ' MBT 0002401 : TFS 0097976 lekf
        ' in VB.NET need declare as Integer for 32bit signed integer
        Dim nMaxTests As Integer

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":TestExecSL1_BeforeTestplanBegin"

        'save the old mouse pointer if this is the first time though a loop
        Select Case muTxSLSimpleState
            Case Is = TxSLSimpleState.running, TxSLSimpleState.paused
                'do nothing
            Case Else
                mnOldMousePointer = Me.Cursor
        End Select

        'change the cursor, for a good indication of progress
        Cursor = System.Windows.Forms.Cursors.WaitCursor

        'Update the local tracking of TxSL state
        muTxSLSimpleState = TxSLSimpleState.running

        If gbTxSLDebugEvents Then System.Diagnostics.Debug.WriteLine("BeforeTestplanBegin:  Count = " & eventArgs.count)
        If gbTxSLTestplanTimeEnabled Then gdtTestplanStartTime = TimeOfDay

        'Disable the failure timer, as we have responded to the run command
        tmrFailure.Enabled = False

        'Keep the the barcode entry field enabled while running
        'This is to allow for catching errant bar codes while running
        'We will simply beep and ignore them.
        'The alternative was to let the "enter" at the end of the bar code be
        'ignored by a inactive control.  It would then be treated like an enter
        'on the key that would have the focus, which was the stop.  The result--stops
        'that you didn't want.
        lblBarCode.Enabled = True
        medBarCode.Enabled = True
        medBarCode.BackColor = System.Drawing.SystemColors.InactiveBorder 'a nice gray

        'Manage configurations
        'set the system status display to running
        rtbSystemStatus.SelectionColor = Color.Cyan
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnTesting)

        'set the operator display
        rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnRunningOpMsg)
        rtbReport.BackColor = System.Drawing.Color.White

        'configure the progress bar
        fraTestplanProgress.Enabled = True
        With prbTestplan
            ' MBT 0002401 : TFS 0097976 lekf
            ' MBT 0002651 : TFS 0105176 lekf , make the ProgressBar display correctly
            nMaxTests = TestExecSL1.Testplan.ExecutableTestCount
            If (nMaxTests) = 0 Then
                .Maximum = 1 'Can't tolerate a max = min
            Else
                .Maximum = nMaxTests
            End If
            ' MBT 0002651 : TFS 0105176 lekf , make the ProgressBar display correctly, need to check the range here
            If (TestExecSL1.Testplan.CurrentTestCount <= .Maximum) Then
                .Value = TestExecSL1.Testplan.CurrentTestCount
            Else
                .Value = .Maximum
            End If
            .Enabled = True
        End With

        'enable the staDescriptions area
        'staDescriptions.Panels(iCURRENTTESTNAME).Enabled = True

        ConfigButtonsToStop()

        'Set the focus depending on the presence of a bar code
        If gbWedgeKeyboardBarCode = True Then
            medBarCode.Focus()
        Else
            cmdStop.Focus()
        End If

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    '***************************************************************************
    Public Sub AfterSuccessfulLogin()

        'a sub that configures most of the controls to be enabled
        'and in an appropriate state after a successful login
        frmLoadTestplan.PopulateLoadTestplanControl()

        'configure the testplan execution buttons
        cmdRun.Enabled = False
        'cmdStep.Enabled = False
        'cmdContinue.Enabled = False
        'cmdPause.Enabled = False
        cmdStop.Enabled = False
        cmdAbort.Enabled = False

        'Configure the bar code areas
        lblBarCode.Enabled = True
        medBarCode.Enabled = True
        medBarCode.BackColor = System.Drawing.Color.White
        cmdLoadTestplan.Enabled = True

        'configure the controls to allow login, load , exit or bar code
        cmdLogin.Enabled = True
        cmdLoadTestplan.Enabled = True
        lblBarCode.Enabled = True

        'Manage configurations
        If gbTxSLSharedIO = True Then frmTxSLSharedIO.AfterLogin()

        'set the system status display to Empty
        fraSystemStatus.Enabled = True
        With rtbSystemStatus
            .SelectionColor = Color.Black
            .Text = LangLookup(modLocalization.txslLangIndex.gnEmpty)
        End With

        'set the operator display
        fraOperatorMessage.Enabled = True
        rtbOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnPleaseLoadTestplan)

        If gbWedgeKeyboardBarCode Then
            medBarCode.Enabled = True
            medBarCode.Focus()
        Else
            cmdLoadTestplan.Enabled = True
            cmdLoadTestplan.Focus()
        End If

        Dim thread As New Thread(AddressOf TcpThread)
        thread.Start()
    End Sub

    ReadOnly _tcpClient As New TcpClient()
    ReadOnly _readyToRunEvent As New AutoResetEvent(False)
    ReadOnly _testDoneEvent As New AutoResetEvent(False)
    Private Sub TcpThread()
        If Not _tcpClient.Connect() Then
            Return
        End If

        SysBox.BackColor = Color.Green
        _readyToRunEvent.WaitOne()

        While True
            If _tcpClient.Check() Then
                FixBox.BackColor = Color.Green
                BeginInvoke(Sub()
                                cmdRun_Click(cmdRun, New System.EventArgs)
                            End Sub)
                _testDoneEvent.WaitOne()
                _testDoneEvent.Reset()
            End If
            Thread.Sleep(500)
        End While
    End Sub

    Private Function YieldCalculation() As Single

        'a short function to hold the yield calculation.
        'It is centralized here, so that custom yield calculations can be changed in
        'one place.
        'This algorithm is very simply, where yield is PassedUUT/TestedUUT
        With TestExecSL1.Testplan.History
            If .TestedUUTCount = 0 Then
                YieldCalculation = 0.0#
            Else
                YieldCalculation = CSng(.PassedUUTCount) / CSng(.TestedUUTCount)
            End If
        End With

    End Function

    Private Sub tmrFailure_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrFailure.Tick
        'A timer that is used to monitor if an attempt at run, stop, abort ....
        'actually happen.  If this timer fires, then the appropriate event that should
        'have happened (like a BeforeTestplanBegin event, as a consequence of sending run)
        'did NOT happen.  We are probably in deep trouble (as TxSL did not evidently respond)
        'but we will attempt to return
        'control to the operator, and let them decide.

        If gbTxSLDebugEvents = True Then System.Diagnostics.Debug.WriteLine("tmrFailure event" & "  " & TimeOfDay)

        Select Case msCommandAttempt
            Case "Run"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup((modLocalization.txslLangIndex.gnRunFailed))
                rtbOperatorMessage.Text = "The Run command timed out, and the Run button has been reenabled."
                ConfigButtonsBeforeRun()
                ConfigButtonsAfterStop()
            Case "Stop"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup((modLocalization.txslLangIndex.gnStopFailed))
                rtbOperatorMessage.Text = "The Stop command timed out, and the Stop button has been reenabled."
                ConfigButtonsToStop()
            Case "Abort"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup((modLocalization.txslLangIndex.gnAbortFailed))
                rtbOperatorMessage.Text = "The Abort command timed out, and the Abort button has been reenabled."
                ConfigButtonsToStop()
            Case "Exit"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = LangLookup((modLocalization.txslLangIndex.gnExitFailed))
                rtbOperatorMessage.Text = "The Exit command timed out, and the Exit button has been reenabled."
                cmdTxSLExit.Enabled = False
            Case "Continue"
                rtbSystemStatus.SelectionColor = Color.Black
                rtbSystemStatus.Text = "Continue Failed"
                rtbOperatorMessage.Text = "The Continue command timed out, and the Continue button has been reenabled."
                ConfigButtonsBeforeRun()
                cmdRun.Text = LangLookup((modLocalization.txslLangIndex.gnContinue))
            Case Else
                MsgBox("Unknown exit flag for tmrFailure event handler", MsgBoxStyle.OkOnly, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)

        End Select
        'Turn off the timer upon it's first event, this makes it a single shot timer.
        tmrFailure.Enabled = False

    End Sub

    Private Sub txtSerialNumber_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerialNumber.TextChanged

        txtSerialNumberChange = True

    End Sub


    Private Sub txtSerialNumber_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSerialNumber.Leave 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":txtSerialNumber_LostFocus"
        If txtSerialNumberChange Then
            UpdateTxSLSerialNumberFields((txtSerialNumber.Text))
            txtSerialNumberChange = False
        End If
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Public Sub UpdateOperatorInterface() 'TxSLErrorTrapSite

        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":UpdateOperatorInterface"
        'This sub should be called from either an AdviseUpdateEvent, or a AfterTestDone Event.
        'In either case, it updates selected fields on the operator interface

        'Check for overflow, and if found, lengthen the progress bar
        'overflow might be encountered if we had looping going on in the system
        ' MBT 0002401 : TFS 0097976 lekf
        ' MBT 0002651 : TFS 0105176 lekf , make the ProgressBar display correctly, the original code is work fine here
        'Dim cMax As Integer
        'cMax = (TestExecSL1.Testplan.CurrentTestCount + TestExecSL1.Testplan.ExecutableTestCount)
        If TestExecSL1.Testplan.CurrentTestCount > prbTestplan.Maximum Then
            prbTestplan.Maximum = TestExecSL1.Testplan.CurrentTestCount + TestExecSL1.Testplan.ExecutableTestCount
        End If

        'update the progress bar
        prbTestplan.Value = TestExecSL1.Testplan.CurrentTestCount

        'Update the test field with the new test name
        'Note that if we are bewteen tests (or on an "other" statement, like a
        'an "if" or "loop", we report a blank name
        'In that case, we put up a string of "Between tests" and exit

        Static TestName As String
        TestName = TestExecSL1.Testplan.CurrentTestExecuting
        If TestName = "" Then
            staDescriptions.Panels(iCURRENTTESTNAME).Text = "Between tests"
        Else
            staDescriptions.Panels(iCURRENTTESTNAME).Text = txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnTest1Param), TestName)
        End If

        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Public Sub UpdateTxSLSerialNumberFields(ByRef SerialNumber As String)
        'A general utility to write to the TxSL system symbol table.
        'Put here as it may be called from multiple places
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":UpdateTxSLSerialNumberFields"
        TestExecSL1.SymbolTables.Item("System").Symbols.Item("SerialNumber").Value = SerialNumber
        Exit Sub

LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)

    End Sub

    Public Sub ConfigControlsAfterLoadFailed()
        'A sub to configure the controls after the failure to load a testplan.

        medBarCode.CtlText = ""

        'Set the controls, depending on the whether a testplan remains loaded
        If TestExecSL1.Testplan.State = TxSLSimpleState.NoTestplan Then
            cmdRun.Enabled = False
            txtSerialNumber.Enabled = False
            cmdSelectVariant.Enabled = False
        Else
            cmdRun.Enabled = True
            txtSerialNumber.Enabled = True
            cmdSelectVariant.Enabled = True
        End If

        'Set the focus appropriately, depending on the presence of a keyboard wedge bar code reader
        'and whether we have errors present

        medBarCode.Enabled = True
        cmdLoadTestplan.Enabled = True
        If frmErrorDialog.ErrorCount = 0 Then
            If gbWedgeKeyboardBarCode = True Then
                medBarCode.Focus()
            Else
                cmdLoadTestplan.Focus()
            End If
        Else '
            'We still have errors present, leave the focus on the error dialog
            'do nothing
        End If

        rtbSystemStatus.SelectionColor = Color.Red
        rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnLoadFailed)
        rtbOperatorMessage.Text = ""


    End Sub


    Public Sub ConfigButtonsAfterStop()
        'a helper sub to centralize the setting buttons after the testplan has returned
        'from a run command
        cmdLogin.Enabled = True
        cmdLoadTestplan.Enabled = True
        cmdSelectVariant.Enabled = True
    End Sub

    Public Sub ConfigButtonsToStop()
        'A helper sub to centralize the setting of buttons to a stop state.
        'set up Execution buttons to the appropriate states
        'cmdPause.Enabled = True
        cmdStop.Enabled = True
        cmdAbort.Enabled = True
    End Sub

    Public Sub Synch()
        'a function to attempt to synch the button states with the state of TxSL.
        'It is necessary because VB and TxSL are operating in two different processes
        'If we halt VB (via debugging or a dialog that VB pops up), TxSL will continue
        'to generate events, that VB will ignore. If we then continue VB, we will be out of synch
        'This sub attempts to read the state of TxSL, and set the button states appropriately

        Select Case TestExecSL1.Testplan.State

            Case HPTestExecSL.TestplanState.NoTestplan
                AfterSuccessfulLogin()
            Case HPTestExecSL.TestplanState.TestplanNotRun
                ConfigButtonsBeforeRun()
            Case HPTestExecSL.TestplanState.TestplanRunning
                ConfigButtonsToStop()
            Case HPTestExecSL.TestplanState.TestplanPaused To HPTestExecSL.TestplanState.TestplanStepPause
                'Call the AfterTestplanPauseEvent with bogus state, and unknown test
                'We will use 99 as the state
                TestExecSL1_AfterTestplanPause(TestExecSL1, New AxHPTestExecSL.DTestExecSLEvents_AfterTestplanPauseEvent(99, "Unknown"))
                TestExecSL1.Testplan.SendReportMessage("Synching up after dialog--possible loss of data in VB reports")
            Case HPTestExecSL.TestplanState.TestplanError To HPTestExecSL.TestplanState.TestplanUnhandledException
                'Call the AfterTestplanStop with a bogus state
                'We will use 99 as the state
                TestExecSL1_AfterTestplanStop(TestExecSL1, New AxHPTestExecSL.DTestExecSLEvents_AfterTestplanStopEvent(99))
                TestExecSL1.Testplan.SendReportMessage("Synching up after dialog--possible loss of data in VB reports")
            Case Else
                MsgBox("Unknown state in sub me.synch", MsgBoxStyle.OkOnly, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
        End Select


    End Sub

    Private Sub fraExecution_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fraExecution.Enter

    End Sub

    Private Sub prbTestplan_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prbTestplan.Click

    End Sub

    Private Sub DebugCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DebugCheckBox.CheckedChanged
        DebugNumericUpDown.Enabled = DebugCheckBox.Checked
    End Sub

    Private Sub DebugNumericUpDown_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DebugNumericUpDown.ValueChanged
        Try
            TestExecSL1.Testplan.Preference.FailCountLimit = DebugNumericUpDown.Value
        Catch ex As AxHost.InvalidActiveXStateException
            ' Not activated yet. Just pass.
        End Try
    End Sub
End Class
