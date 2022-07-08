Option Strict Off
Option Explicit On
Friend Class frmSelectVariant
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
    Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents lblSelectANewVariant As System.Windows.Forms.Label
	Public WithEvents Line1 As System.Windows.Forms.Label
	Public WithEvents lblCurrentVariant As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
    Friend WithEvents lsvVariantDetails As System.Windows.Forms.ListView
    Friend WithEvents rtbCurrentVariant As System.Windows.Forms.RichTextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOK = New System.Windows.Forms.Button
        Me.lblSelectANewVariant = New System.Windows.Forms.Label
        Me.Line1 = New System.Windows.Forms.Label
        Me.lblCurrentVariant = New System.Windows.Forms.Label
        Me.lsvVariantDetails = New System.Windows.Forms.ListView
        Me.rtbCurrentVariant = New System.Windows.Forms.RichTextBox
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(241, 199)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(89, 33)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdOK
        '
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.Enabled = False
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Location = New System.Drawing.Point(241, 148)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.Size = New System.Drawing.Size(89, 33)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "OK"
        '
        'lblSelectANewVariant
        '
        Me.lblSelectANewVariant.BackColor = System.Drawing.SystemColors.Control
        Me.lblSelectANewVariant.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSelectANewVariant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectANewVariant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSelectANewVariant.Location = New System.Drawing.Point(37, 120)
        Me.lblSelectANewVariant.Name = "lblSelectANewVariant"
        Me.lblSelectANewVariant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSelectANewVariant.Size = New System.Drawing.Size(193, 17)
        Me.lblSelectANewVariant.TabIndex = 1
        Me.lblSelectANewVariant.Text = "Select a new Variant"
        '
        'Line1
        '
        Me.Line1.BackColor = System.Drawing.SystemColors.WindowText
        Me.Line1.Location = New System.Drawing.Point(36, 104)
        Me.Line1.Name = "Line1"
        Me.Line1.Size = New System.Drawing.Size(304, 1)
        Me.Line1.TabIndex = 5
        '
        'lblCurrentVariant
        '
        Me.lblCurrentVariant.BackColor = System.Drawing.SystemColors.Control
        Me.lblCurrentVariant.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCurrentVariant.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentVariant.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurrentVariant.Location = New System.Drawing.Point(36, 45)
        Me.lblCurrentVariant.Name = "lblCurrentVariant"
        Me.lblCurrentVariant.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCurrentVariant.Size = New System.Drawing.Size(284, 17)
        Me.lblCurrentVariant.TabIndex = 0
        Me.lblCurrentVariant.Text = "Current Variant"
        '
        'lsvVariantDetails
        '
        Me.lsvVariantDetails.Location = New System.Drawing.Point(40, 144)
        Me.lsvVariantDetails.MultiSelect = False
        Me.lsvVariantDetails.Name = "lsvVariantDetails"
        Me.lsvVariantDetails.Size = New System.Drawing.Size(192, 88)
        Me.lsvVariantDetails.TabIndex = 6
        '
        'rtbCurrentVariant
        '
        Me.rtbCurrentVariant.Location = New System.Drawing.Point(40, 64)
        Me.rtbCurrentVariant.Name = "rtbCurrentVariant"
        Me.rtbCurrentVariant.Size = New System.Drawing.Size(192, 24)
        Me.rtbCurrentVariant.TabIndex = 7
        Me.rtbCurrentVariant.Text = ""
        '
        'frmSelectVariant
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(404, 273)
        Me.Controls.Add(Me.rtbCurrentVariant)
        Me.Controls.Add(Me.lsvVariantDetails)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdOK)
        Me.Controls.Add(Me.lblSelectANewVariant)
        Me.Controls.Add(Me.Line1)
        Me.Controls.Add(Me.lblCurrentVariant)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(162, 205)
        Me.Name = "frmSelectVariant"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Variant"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmSelectVariant
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmSelectVariant
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmSelectVariant()
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
	'frmSelectVariant.frm
	'frmSelectVariant.frx
	'
	'The frmSelectVariant.frm form handles the task of selecting
	'a testplan variant.  Note that selection of a variant occurs
	'automatically when a testplan is first loaded.  In that case,
	'the "default" variant, as specified in the Testplan Registry of
	'the tstexcsl.ini file is selected.  Selection of non default
	'variants is done by clicking the SelectVariant button on the
	'main form, which in turn, will pop up the frmSelectVariant.frm
	'and use the logic contained within it.
	'*****************************************************************************
	
	
	Private LocalSource As String
	
	Public Sub PopulateVariantsBox() 'TxSLErrorTrapSite
		'A sub to populate the list view box with elements, following the load
		'of a testplan.
		On Error GoTo LocalErrorHandler
		LocalSource = ":" & Me.Name & ":AfterTestplanLoad"
		
		
		
        lsvVariantDetails.Items.Clear()
		
		'charge the listview box with data.
		'Note that the TestplanVariants collection is a collection of strings, and NOT objects.
		'For this reason, we can not use the For Each approach of traversing the collection,
		'but instead use the For i = 1 to count approach.
        Dim itmX As ListViewItem
		Dim i As Short
        For i = 1 To frmMain.TestExecSL1.Testplan.TestplanVariants.Count
            itmX = lsvVariantDetails.Items.Add(frmMain.TestExecSL1.Testplan.TestplanVariants.Item(i))
        Next i

        Exit Sub
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Hide()
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        SelectVariant(lsvVariantDetails.SelectedItems(0).Text)
		Me.Hide()
	End Sub
	
	Private Sub frmSelectVariant_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		lsvVariantDetails.Focus()
	End Sub
	
	'Private Sub lsvVariantDetails_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
	'  'code to allow the listview to be sorted, by column clicks.
	'  lsvVariantDetails.SortKey = ColumnHeader.Index - 1
	'  lsvVariantDetails.Sorted = True
	'  If lsvVariantDetails.SortOrder = lvwAscending Then
	'    lsvVariantDetails.SortOrder = lvwDescending
	'  Else
	'    lsvVariantDetails.SortOrder = lvwAscending
	'  End If
	'
	'End Sub

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

    Public Function SelectVariant(ByRef VariantName As String) As Boolean 'TxSLErrorTrapSite
        On Error GoTo LocalErrorHandler
        LocalSource = ":" & Me.Name & ":SelectVariant"

        'A helper function to call the select variant function, and, if
        'successful, update the SelectVariant box, as well as the status bar
        'on the main form.  This indirection is employed in case the variant is
        'being selected programmatically, and not from the UI.

        Dim itmFound As ListViewItem
        If frmMain.TestExecSL1.Testplan.SelectVariant(VariantName) <> True Then
            'It was not successfully selected.  Present an error box, update boxes with "" and return false
            MsgBox(txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUnsuccessfulVariantSelect), VariantName), , System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
            rtbCurrentVariant.Text = ""
            frmLoadTestplan.txtCurrentVariant.Text = ""
            'frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).Enabled = True
            frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnVariant) & ": " & ""
            SelectVariant = False
        Else
            'It was selected successfully
            itmFound = lsvVariantDetails.Items(FindListViewItem(lsvVariantDetails, VariantName, -1))
            itmFound.EnsureVisible()
            itmFound.Selected = True
            rtbCurrentVariant.Text = itmFound.Text
            frmLoadTestplan.txtCurrentVariant.Text = itmFound.Text
            'frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).Enabled = True
            frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnVariant) & ": " & Me.rtbCurrentVariant.Text
            SelectVariant = True
        End If

        cmdCancel.Enabled = True
        cmdOK.Enabled = True

        frmMain.cmdSelectVariant.Enabled = True

        Exit Function
LocalErrorHandler:
        frmErrorDialog.ErrorHandler(Err.Number, Err.Source & LocalSource, Err.Description)
    End Function

    Private Sub lsvVariantDetails_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lsvVariantDetails.DoubleClick
        SelectVariant(lsvVariantDetails.SelectedItems(0).Text)
        Me.Hide()
    End Sub

    Public Sub AfterTestplanUnload()
        rtbCurrentVariant.Text = ""
        lsvVariantDetails.Items.Clear()
        lsvVariantDetails.Enabled = False
        cmdOK.Enabled = False
    End Sub

    Public Function VariantIsValid(ByRef VariantName As String) As Boolean
        'A function to confirm that the VariantName (probably the default variant)
        'is contained in the collection of variant names that were populated
        'in the lsvVariantDetails box.

        Dim itmFound As ListViewItem
        itmFound = lsvVariantDetails.Items(FindListViewItem(lsvVariantDetails, VariantName, -1))
        If itmFound Is Nothing Then
            MsgBox(txslFormatString1(LangLookup(modLocalization.txslLangIndex.gnUnknownVariantOpmsg), VariantName), , System.Reflection.Assembly.GetExecutingAssembly.GetName.Name)
            VariantIsValid = False
        Else
            VariantIsValid = True
        End If

    End Function

    Private Sub lsvVariantDetails_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lsvVariantDetails.KeyDown
        Select Case e.KeyCode
            Case System.Windows.Forms.Keys.Up, System.Windows.Forms.Keys.Down
                'do nothing, this is allowed behavior
            Case System.Windows.Forms.Keys.Return
                cmdOK_Click(cmdOK, New System.EventArgs)
            Case Else
                Beep()
        End Select
    End Sub

End Class