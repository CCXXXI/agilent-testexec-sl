Option Strict Off
Option Explicit On
Friend Class frmLoadTestplan
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
    Public WithEvents tmrRefresh As System.Windows.Forms.Timer
    Public WithEvents cmdOk As System.Windows.Forms.Button
	Public WithEvents txtCurrentTestplanName As System.Windows.Forms.TextBox
	Public WithEvents txtCurrentVariant As System.Windows.Forms.TextBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents lblName As System.Windows.Forms.Label
	Public WithEvents lblPath As System.Windows.Forms.Label
	Public WithEvents Line1 As System.Windows.Forms.Label
	Public WithEvents lblCurrentTestplan As System.Windows.Forms.Label
	Public WithEvents lblTestplanDescription As System.Windows.Forms.Label
	Public WithEvents lblCurrentVariant As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Public WithEvents lsvTestplanDetails As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imlTestplan As System.Windows.Forms.ImageList
    Friend WithEvents rtbPath As System.Windows.Forms.RichTextBox
    Friend WithEvents rtbTestplanDescription As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLoadTestplan))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCurrentTestplanName = New System.Windows.Forms.TextBox()
        Me.txtCurrentVariant = New System.Windows.Forms.TextBox()
        Me.tmrRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.Line1 = New System.Windows.Forms.Label()
        Me.lblCurrentTestplan = New System.Windows.Forms.Label()
        Me.lblTestplanDescription = New System.Windows.Forms.Label()
        Me.lblCurrentVariant = New System.Windows.Forms.Label()
        Me.lsvTestplanDetails = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.imlTestplan = New System.Windows.Forms.ImageList(Me.components)
        Me.rtbPath = New System.Windows.Forms.RichTextBox()
        Me.rtbTestplanDescription = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'txtCurrentTestplanName
        '
        Me.txtCurrentTestplanName.AcceptsReturn = True
        Me.txtCurrentTestplanName.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrentTestplanName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrentTestplanName.Enabled = False
        Me.txtCurrentTestplanName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurrentTestplanName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrentTestplanName.Location = New System.Drawing.Point(28, 53)
        Me.txtCurrentTestplanName.MaxLength = 0
        Me.txtCurrentTestplanName.Name = "txtCurrentTestplanName"
        Me.txtCurrentTestplanName.ReadOnly = True
        Me.txtCurrentTestplanName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrentTestplanName.Size = New System.Drawing.Size(164, 25)
        Me.txtCurrentTestplanName.TabIndex = 0
        Me.txtCurrentTestplanName.TabStop = False
        Me.ToolTip1.SetToolTip(Me.txtCurrentTestplanName, "The name of the current testplan")
        '
        'txtCurrentVariant
        '
        Me.txtCurrentVariant.AcceptsReturn = True
        Me.txtCurrentVariant.BackColor = System.Drawing.SystemColors.Window
        Me.txtCurrentVariant.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCurrentVariant.Enabled = False
        Me.txtCurrentVariant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCurrentVariant.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCurrentVariant.Location = New System.Drawing.Point(28, 122)
        Me.txtCurrentVariant.MaxLength = 0
        Me.txtCurrentVariant.Name = "txtCurrentVariant"
        Me.txtCurrentVariant.ReadOnly = True
        Me.txtCurrentVariant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCurrentVariant.Size = New System.Drawing.Size(164, 24)
        Me.txtCurrentVariant.TabIndex = 1
        Me.txtCurrentVariant.TabStop = False
        Me.ToolTip1.SetToolTip(Me.txtCurrentVariant, "The current testplan variant")
        '
        'tmrRefresh
        '
        Me.tmrRefresh.Interval = 200
        '
        'cmdOk
        '
        Me.cmdOk.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOk.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOk.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOk.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOk.Location = New System.Drawing.Point(422, 197)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOk.Size = New System.Drawing.Size(117, 50)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "OK"
        Me.cmdOk.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(422, 276)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(117, 50)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'lblName
        '
        Me.lblName.BackColor = System.Drawing.SystemColors.Control
        Me.lblName.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblName.Location = New System.Drawing.Point(29, 177)
        Me.lblName.Name = "lblName"
        Me.lblName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblName.Size = New System.Drawing.Size(289, 21)
        Me.lblName.TabIndex = 10
        Me.lblName.Text = "Name"
        '
        'lblPath
        '
        Me.lblPath.BackColor = System.Drawing.SystemColors.Control
        Me.lblPath.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPath.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPath.Location = New System.Drawing.Point(29, 335)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPath.Size = New System.Drawing.Size(174, 21)
        Me.lblPath.TabIndex = 8
        Me.lblPath.Text = "Path"
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(28, 160)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(470, 1)
        Me.Line1.TabIndex = 13
        '
        'lblCurrentTestplan
        '
        Me.lblCurrentTestplan.BackColor = System.Drawing.SystemColors.Control
        Me.lblCurrentTestplan.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCurrentTestplan.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentTestplan.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurrentTestplan.Location = New System.Drawing.Point(28, 20)
        Me.lblCurrentTestplan.Name = "lblCurrentTestplan"
        Me.lblCurrentTestplan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCurrentTestplan.Size = New System.Drawing.Size(204, 21)
        Me.lblCurrentTestplan.TabIndex = 6
        Me.lblCurrentTestplan.Text = "Current Testplan"
        '
        'lblTestplanDescription
        '
        Me.lblTestplanDescription.BackColor = System.Drawing.SystemColors.Control
        Me.lblTestplanDescription.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTestplanDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestplanDescription.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTestplanDescription.Location = New System.Drawing.Point(30, 422)
        Me.lblTestplanDescription.Name = "lblTestplanDescription"
        Me.lblTestplanDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTestplanDescription.Size = New System.Drawing.Size(328, 21)
        Me.lblTestplanDescription.TabIndex = 5
        Me.lblTestplanDescription.Text = "Testplan Description"
        '
        'lblCurrentVariant
        '
        Me.lblCurrentVariant.BackColor = System.Drawing.SystemColors.Control
        Me.lblCurrentVariant.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCurrentVariant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentVariant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurrentVariant.Location = New System.Drawing.Point(28, 89)
        Me.lblCurrentVariant.Name = "lblCurrentVariant"
        Me.lblCurrentVariant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCurrentVariant.Size = New System.Drawing.Size(171, 21)
        Me.lblCurrentVariant.TabIndex = 4
        Me.lblCurrentVariant.Text = "Current Variant"
        '
        'lsvTestplanDetails
        '
        Me.lsvTestplanDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
        Me.lsvTestplanDetails.LargeImageList = Me.imlTestplan
        Me.lsvTestplanDetails.Location = New System.Drawing.Point(29, 207)
        Me.lsvTestplanDetails.MultiSelect = False
        Me.lsvTestplanDetails.Name = "lsvTestplanDetails"
        Me.lsvTestplanDetails.Size = New System.Drawing.Size(365, 118)
        Me.lsvTestplanDetails.SmallImageList = Me.imlTestplan
        Me.lsvTestplanDetails.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lsvTestplanDetails.TabIndex = 14
        Me.lsvTestplanDetails.UseCompatibleStateImageBehavior = False
        Me.lsvTestplanDetails.View = System.Windows.Forms.View.List
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        Me.ColumnHeader1.Width = 180
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Default Variant"
        Me.ColumnHeader2.Width = 180
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Path"
        Me.ColumnHeader3.Width = 180
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Description"
        Me.ColumnHeader4.Width = 180
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Short Description"
        Me.ColumnHeader5.Width = 180
        '
        'imlTestplan
        '
        Me.imlTestplan.ImageStream = CType(resources.GetObject("imlTestplan.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imlTestplan.TransparentColor = System.Drawing.Color.Transparent
        Me.imlTestplan.Images.SetKeyName(0, "")
        '
        'rtbPath
        '
        Me.rtbPath.Location = New System.Drawing.Point(29, 364)
        Me.rtbPath.Name = "rtbPath"
        Me.rtbPath.Size = New System.Drawing.Size(365, 50)
        Me.rtbPath.TabIndex = 15
        Me.rtbPath.Text = ""
        '
        'rtbTestplanDescription
        '
        Me.rtbTestplanDescription.Location = New System.Drawing.Point(29, 443)
        Me.rtbTestplanDescription.Name = "rtbTestplanDescription"
        Me.rtbTestplanDescription.Size = New System.Drawing.Size(365, 148)
        Me.rtbTestplanDescription.TabIndex = 16
        Me.rtbTestplanDescription.Text = ""
        '
        'frmLoadTestplan
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(616, 613)
        Me.Controls.Add(Me.rtbTestplanDescription)
        Me.Controls.Add(Me.rtbPath)
        Me.Controls.Add(Me.lsvTestplanDetails)
        Me.Controls.Add(Me.cmdOk)
        Me.Controls.Add(Me.txtCurrentTestplanName)
        Me.Controls.Add(Me.txtCurrentVariant)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.lblCurrentTestplan)
        Me.Controls.Add(Me.lblTestplanDescription)
        Me.Controls.Add(Me.lblCurrentVariant)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(109, 49)
        Me.Name = "frmLoadTestplan"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Load Testplan"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmLoadTestplan
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmLoadTestplan
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmLoadTestplan()
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
	'frmLoadTestplan.frm
	'frmLoadTestplan.frx
	'
	'These files support the loading of testplans.  They are designed around
	'reading the Testplan Registry section of the tstexcsl.ini file,
	'and will only load files that have been specified in that file.
	'See the txtexcsl.ini file for more information.
	'*****************************************************************************
	
	'Constants for use in loading the testplan listview box
	'They are useful to keep the code self documenting
	Private Const iDEFAULT_VARIANT As Short = 1
	Private Const iPATH As Short = 2
	Private Const iDESCRIPTION As Short = 3
	Private Const iSHORTDESCRIPTION As Short = 4
	Private LocalSource As String
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'Upon pushing the OK button, we will get the selected item from the listview box,
		'find the path associated with that item, and pass that pass to the LoadTestplan
		'sub, contained later on this form.
		
        If LoadTestplan(lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text, lsvTestplanDetails.SelectedItems(0).SubItems(iDEFAULT_VARIANT).Text) Then
            tmrRefresh.Enabled = False
            Hide() 'the load was successful
        Else
            lsvTestplanDetails.Focus()
        End If
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		tmrRefresh.Enabled = False
		Hide()
	End Sub
	
	Private Sub frmLoadTestplan_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'tmrRefresh.Enabled = True
		lsvTestplanDetails.Enabled = True
		lsvTestplanDetails.Focus()
		
	End Sub
	
	Private Sub frmLoadTestplan_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		PopulateLoadTestplanControl()
		tmrRefresh.Enabled = True
		
    End Sub

    Private Sub lsvTestplanDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lsvTestplanDetails.Click
        cmdOk.Enabled = True
        rtbTestplanDescription.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iDESCRIPTION).Text
        rtbPath.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text
    End Sub

    Private Sub lsvTestplanDetails_ColumnClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lsvTestplanDetails.ColumnClick
        'code to allow the listview to be sorted, by column clicks.
        'only applies if the report view is selected, which is not the standard presentation

        Dim sort As SortOrder = lsvTestplanDetails.Sorting
        If sort = SortOrder.Ascending Then
            sort = SortOrder.Descending
        Else
            sort = SortOrder.Ascending
        End If

        ' Set the ListViewItemSorter property to a new ListViewItemComparer 
        ' object. Setting this property immediately sorts the 
        ' ListView using the ListViewItemComparer object.
        lsvTestplanDetails.ListViewItemSorter = New ListViewItemComparer(e.Column)
        lsvTestplanDetails.Sorting = sort
    End Sub

    Public Sub PopulateLoadTestplanControl() 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":PopulateLoadTestplanControl"

        lsvTestplanDetails.Items.Clear()

        'charge the listview box with data.
        Dim itmX As ListViewItem
        'Note that we dimension as a generic object.
        'If we had set a reference to olecore.dll, we could have dim'ed it
        'as a registered testplan.

        Dim RegisteredTestplanInstance As HPTestCoreRuntime.RegisteredTestplan
        Dim NumRegisteredTestplans As Integer = frmMain.TestExecSL1.RegisteredTestplans.Count
        For I As Integer = 1 To NumRegisteredTestplans
            RegisteredTestplanInstance = frmMain.TestExecSL1.RegisteredTestplans.Item(I)
            'we've already checked to make sure that there are testplans
            'For Each RegisteredTestplanInstance In frmMain.TestExecSL1.RegisteredTestplans
            Dim tpaName As String = RegisteredTestplanInstance.Name
            itmX = lsvTestplanDetails.Items.Add(RegisteredTestplanInstance.Name, 0)
            'Note that the subindex must be an integer, and can not be a string or string index

            itmX.SubItems.Add(RegisteredTestplanInstance.DefaultVariant)
            itmX.SubItems.Add(RegisteredTestplanInstance.Path)
            itmX.SubItems.Add(RegisteredTestplanInstance.Description)
            itmX.SubItems.Add(RegisteredTestplanInstance.ShortDescription)

            'Next RegisteredTestplanInstance
        Next I

        cmdOk.Enabled = True
        cmdCancel.Enabled = True

        Exit Sub
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Sub
    '***************************************************************************
    Public Function LoadTestplan(ByRef Path As String, ByRef TpVariant As String) As Boolean 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":LoadTestplan"
        'a helper function that calls the TxSL load testplan
        'method, and then adds error checking.  Put here in one place so
        'that it is callable from both human and bar code operation.
        Dim LoadTestplanResult As Boolean

        'Disable the controls, both to give a visual indicator that we are working
        'and so that we can not get any more inputs, until after we have returned.
        cmdOk.Enabled = False
        cmdCancel.Enabled = False
        lsvTestplanDetails.Enabled = False

        frmMain.rtbSystemStatus.SelectionColor = Color.Black
        frmMain.rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnLoadingTestplan)

        'Here's the actual call to load the testplan
        LoadTestplanResult = frmMain.TestExecSL1.LoadTestplan(Path)
        'If a new testplan was actually loaded, then the AfterTestplanLoad event
        'will have fired, and the event handling takes place here

        If LoadTestplanResult = False Then 'the load failed
            'Update this form
            frmMain.rtbSystemStatus.SelectionColor = Color.Red
            frmMain.rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnLoadFailed)
            LoadTestplan = False

        Else 'The load has apparently succeeded, either by loading a new testplan, or the testplan was
            'already loaded, and we returned immediately.
            'Prior to LoadTestplan returning, the AfterTestplanLoad event will have fired.  In the
            'event handler for AfterTestplanLoad, we will have done additional checks on the validity of the
            'testplan that we have loaded.  These involve checking the path and default variant for validity,
            'and if both ok, then attempting to select the default variant.  A summary of the results of
            'those tests is held in gbAfterTPLoadChecksOkay, which will be true if everything checks out
            'If it is true, then we believe that we really had a successful load, and return true to this
            'function

            If frmMain.gbAfterTPLoadChecksOkay Then
                'This next code loads the variant of the item we have selected
                'Note that if a testplan is in memory, we don't actually reload it
                'Hence, selection of new variants needs to be done here, and not in AfterTestplanLoad event.
                'Failure to be able to select the variant is viewed is viewed as a failure.
                If frmSelectVariant.SelectVariant(TpVariant) Then
                    frmMain.rtbSystemStatus.SelectionColor = Color.Black
                    frmMain.rtbSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnReady)
                    LoadTestplan = True
                Else
                    LoadTestplan = False
                    frmMain.TestExecSL1.UnloadTestplan()
                End If
            Else
                'One of the checks associated with the load failed.
                'See frmMain.testexecsl1.aftertestplanload
                LoadTestplan = False
                frmMain.TestExecSL1.UnloadTestplan()
            End If

        End If

        'Re-enable the controls, to allow a different entry
        Cursor = Cursors.Default
        cmdCancel.Enabled = True
        cmdOk.Enabled = True
        lsvTestplanDetails.Enabled = True

        Exit Function
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Function

    Private Sub lsvTestplanDetails_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lsvTestplanDetails.DoubleClick
        cmdOk.Enabled = True
        Dim tpaPath As String = lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text
        Dim tpaVariant As String = lsvTestplanDetails.SelectedItems(0).SubItems(iDEFAULT_VARIANT).Text

        If LoadTestplan(tpaPath, tpaVariant) Then
            tmrRefresh.Enabled = False
            Me.Hide()
        Else
            lsvTestplanDetails.Focus()
        End If

    End Sub

    Public Sub AfterTestplanUnload()
        txtCurrentTestplanName.Text = ""
        txtCurrentVariant.Text = ""
        rtbTestplanDescription.Text = ""
        cmdOk.Enabled = False

    End Sub

    Public Sub BeforeTestplanLoad()
        'change the enable on these buttons, to give a visual indication
        'that we are loading the testplan
        cmdOk.Enabled = False
        cmdCancel.Enabled = False
        'similarly, change the mouse pointer
        Cursor = System.Windows.Forms.Cursors.WaitCursor
    End Sub

    Public Sub StartSelection()
        'a sub that is called by the Load Testplan button on frmMain.
        'This sub will initialize the frmLoadTestplan buttons to the correct state.
        'If no testplan has been previously selected, it will set the focus to
        'the first tesplan

        cmdCancel.Enabled = True
        cmdOk.Enabled = True

        'similarly, change the mouse pointer
        Cursor = Cursors.Default

        Show()
        'This code looks at the contents of the current testplan name.
        'If it is "", then the focus is set to the first testplan, and the description
        'boxes updated as appropriate.  If not, it makes sure that the correct one
        'is visible

        If txtCurrentTestplanName.Text = "" Then
            lsvTestplanDetails.Items(0).Selected = True
            lsvTestplanDetails.SelectedItems(0).EnsureVisible()
            rtbTestplanDescription.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iDESCRIPTION).Text
            rtbPath.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text
        Else 'Else, make sure that the selected item is visible.
            'The following code finds the items indicated by the name that has been kept
            'in the CurrentTestplanName control.  It selects it, and then makes it visible.
            'lsvTestplanDetails.SelectedItem = lsvTestplanDetails.FindItem(txtCurrentTestplanName.Text, mscomctl.ListFindItemWhereConstants.lvwText, 1, mscomctl.ListFindItemHowConstants.lvwWhole)
            Dim index As Integer = FindListViewItem(lsvTestplanDetails, txtCurrentTestplanName.Text, -1)
            lsvTestplanDetails.Items(index).Selected = True
            lsvTestplanDetails.SelectedItems(0).EnsureVisible()
        End If

    End Sub

    Private Sub lsvTestplanDetails_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lsvTestplanDetails.KeyDown
        'This sub allows a return key to start the load process
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Return
                cmdOK_Click(cmdOk, New System.EventArgs)
            Case Else
                Beep()
        End Select
    End Sub

    Public Function FindListViewItem(ByRef list As ListView, ByVal item As String, ByVal column As Integer) As Integer
        Dim text As String
        Dim numItems As Integer = list.Items.Count
        For i As Integer = 0 To numItems - 1
            If column = -1 Then
                text = list.Items(i).Text()
            Else
                text = list.Items(i).SubItems(column).Text()
            End If

            If item.CompareTo(text) = 0 Then
                FindListViewItem = i
                Exit Function
            End If
        Next

        FindListViewItem = -1
    End Function

    Public Function PathIsValid(ByRef Path As String) As Boolean
        'This uses the lsvTestplanDetails control as a database
        'It checks if the path that was returned by the AfterTestplanLoad event
        'is a valid path, ie. contained in the data base that we read
        'from the .ini file
        Dim itmX As ListViewItem = Nothing

        'itmX = lsvTestplanDetails.FindItem(Path, mscomctl.ListFindItemWhereConstants.lvwSubItem, 1, mscomctl.ListFindItemHowConstants.lvwWhole)
        Dim index As Integer = FindListViewItem(lsvTestplanDetails, Path, iPATH)
        If index <> -1 Then
            itmX = lsvTestplanDetails.Items(index)
        End If
        'if the testplan is not found in the lsv box (which is a data base of registered testplans
        'then put up a dialog, and exit the sub

        If itmX Is Nothing Then
            MsgBox(txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnLoadofUnknowntestplan), Path), , System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
            PathIsValid = False
            Exit Function
        Else
            PathIsValid = True
        End If

    End Function

    Public Function RetrieveDefaultVariant(ByRef Path As String) As String
        'This uses the lsvTestplanDetails control as a database
        'Using the path as an index, it returns the default variant for that path
        Dim itmX As ListViewItem = Nothing
        Dim index As Integer = FindListViewItem(lsvTestplanDetails, Path, iPATH)
        If index <> -1 Then
            itmX = lsvTestplanDetails.Items(index)
        End If
        'if the testplan is not found in the lsv box (which is a data base of registered testplans
        'then put up a dialog, and exit the sub
        If itmX Is Nothing Then
            MsgBox("Can not find a default variant for the testplan at " & Path, MsgBoxStyle.OKOnly, System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
            RetrieveDefaultVariant = ""
        End If
        RetrieveDefaultVariant = itmX.SubItems(iDEFAULT_VARIANT).Text
    End Function

    Public Sub UpdateCaptions(ByRef Path As String)
        'This uses the lsvTestplanDetails control as a database, where
        'paths and default variants can be related back to testplan names.
        'This sub updates this form, knowing that a load has in fact been completed.

        Dim itmX As ListViewItem
        Dim index As Integer = FindListViewItem(lsvTestplanDetails, Path, iPATH)
        itmX = lsvTestplanDetails.Items(index)

        'On this form
        itmX.EnsureVisible()
        itmX.Selected = True
        cmdOk.Enabled = True
        txtCurrentTestplanName.Text = itmX.Text

    End Sub


    Private Sub tmrRefresh_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles tmrRefresh.Tick
        'update the description boxes every so often
        'To minimize flicker, only perform the update if the text is different.

        'This code is frequently the first code that talks to the TxSL control
        'If you are getting a bug here, it is probably due to txsl being installed
        'in a location that we can't find.
        If (rtbTestplanDescription.Text <> lsvTestplanDetails.SelectedItems(0).SubItems(iDESCRIPTION).Text) Then
            rtbTestplanDescription.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iDESCRIPTION).Text
        End If

        If (rtbPath.Text <> lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text) Then
            rtbPath.Text = lsvTestplanDetails.SelectedItems(0).SubItems(iPATH).Text
        End If

    End Sub


End Class