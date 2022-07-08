Option Strict Off
Option Explicit On
Module modAppSpecific
	'*****************************************************************************
	'modAppSpecific.bas
	'
	'The file AppSpecific.bas file holds information specific to a particular
	'application.  From a localization standpoint, it holds the code that
	'relates the localization strings (defined in modLocalization.bas) to the
	'specific controls in the application.  It also holds an example of how
	'to programmatically change the fonts used on the controls.
	'The file modAppSpedific.bas would also be a convenient place to put
	'other "utility" functions.
	'*****************************************************************************
	
	
	Public Sub UpdateControlCaptions()
		'Updates the control captions to a new language
		'It is used with the resources defined in modLocalization.bas
		'This sub is positioned here (as opposed in modLocalization.bas), so that
		'modLocalization.bas can be application independent.
		'The code in this sub is very specific to the application.
		
		frmLoadTestplan.Text = LangLookup(modLocalization.txslLangIndex.gnLoadTestplan)
		frmLoadTestplan.cmdOK.Text = LangLookup(modLocalization.txslLangIndex.gnOK)
		frmLoadTestplan.cmdCancel.Text = LangLookup(modLocalization.txslLangIndex.gnCancel)
		frmLoadTestplan.lblPath.Text = LangLookup(modLocalization.txslLangIndex.gnPath)
		frmLoadTestplan.lblName.Text = LangLookup(modLocalization.txslLangIndex.gnName)
        frmLoadTestplan.lsvTestplanDetails.Columns(0).Text = LangLookup(modLocalization.txslLangIndex.gnName)
        frmLoadTestplan.lsvTestplanDetails.Columns(1).Text = LangLookup(modLocalization.txslLangIndex.gnDefaultVariant)
        frmLoadTestplan.lsvTestplanDetails.Columns(2).Text = LangLookup(modLocalization.txslLangIndex.gnPath)
        frmLoadTestplan.lsvTestplanDetails.Columns(3).Text = LangLookup(modLocalization.txslLangIndex.gnDescription)
        frmLoadTestplan.lsvTestplanDetails.Columns(4).Text = LangLookup(modLocalization.txslLangIndex.gnShortDescription)
		frmLoadTestplan.lblCurrentTestplan.Text = LangLookup(modLocalization.txslLangIndex.gnCurrentTestplan)
		frmLoadTestplan.lblTestplanDescription.Text = LangLookup(modLocalization.txslLangIndex.gnTestplanDescription)
		frmLoadTestplan.lblCurrentVariant.Text = LangLookup(modLocalization.txslLangIndex.gnCurrentVariant)
        frmLogin.Text = LangLookup(modLocalization.txslLangIndex.gnLogIn) & "  Version: " & CStr(System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileVersion) 'NoLocalize& App.Revision & App.Major & App.Minor
		frmLogin.cmdCancel.Text = LangLookup(modLocalization.txslLangIndex.gnCancel)
		frmLogin.cmdOK.Text = LangLookup(modLocalization.txslLangIndex.gnOK)
		frmLogin.lblPassword.Text = LangLookup(modLocalization.txslLangIndex.gnPassword)
		frmLogin.lblName.Text = LangLookup(modLocalization.txslLangIndex.gnName)
		frmMain.Text = LangLookup(modLocalization.txslLangIndex.gnAgilentTechnologiesTestExecSLOperatorInterface)
		frmMain.fraTestplanConfiguration.Text = LangLookup(modLocalization.txslLangIndex.gnTestplanConfiguration)
		frmMain.chkReportPassedTests.Text = LangLookup(modLocalization.txslLangIndex.gnReportPassedTests)
		frmMain.chkReportFailedTests.Text = LangLookup(modLocalization.txslLangIndex.gnReportFailedTests1)
		frmMain.chkReportException.Text = LangLookup(modLocalization.txslLangIndex.gnReportExceptions1)
		frmMain.chkShowReport.Text = LangLookup(modLocalization.txslLangIndex.gnShowReport1)
        frmMain.staStats.Panels(frmMain.iPASSED).Text = LangLookup(modLocalization.txslLangIndex.gnPassed)
        frmMain.staStats.Panels(frmMain.iPASSED).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnUUTPassedTip)
        frmMain.staStats.Panels(frmMain.iFAILED).Text = LangLookup(modLocalization.txslLangIndex.gnFailed)
        frmMain.staStats.Panels(frmMain.iFAILED).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnUUTFailedTip)
        frmMain.staStats.Panels(frmMain.iTOTALTESTED).Text = LangLookup(modLocalization.txslLangIndex.gnTotalTested)
        frmMain.staStats.Panels(frmMain.iTOTALTESTED).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnTotalTestedTip)
        frmMain.staStats.Panels(frmMain.iYIELD).Text = LangLookup(modLocalization.txslLangIndex.gnYieldTip)
        frmMain.staStats.Panels(frmMain.iYIELD).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnYieldTip)
        frmMain.staStats.Panels(frmMain.iSINCE).Text = LangLookup(modLocalization.txslLangIndex.gnSince)
        frmMain.staStats.Panels(frmMain.iSINCE).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnSinceTip)
        frmMain.staStats.Panels(frmMain.iSINCE).Text = LangLookup(modLocalization.txslLangIndex.gnSince)
		frmMain.fraTxSLConfiguration.Text = LangLookup(modLocalization.txslLangIndex.gnTxSLConfiguration)
		frmMain.cmdSelectVariant.Text = LangLookup(modLocalization.txslLangIndex.gnSelectVariant1)
		frmMain.cmdLoadTestplan.Text = LangLookup(modLocalization.txslLangIndex.gnLoadTestplan1)
		frmMain.cmdTxSLExit.Text = LangLookup(modLocalization.txslLangIndex.gnExit1)
		frmMain.cmdLogin.Text = LangLookup(modLocalization.txslLangIndex.gnLogin1)
		frmMain.lblBarCode.Text = LangLookup(modLocalization.txslLangIndex.gnBarCode)
        frmMain.staDescriptions.Panels(frmMain.iUUTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnUUT)
        frmMain.staDescriptions.Panels(frmMain.iUUTNAME).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnUUTNameTip)
        frmMain.staDescriptions.Panels(frmMain.iTESTPLANNAME).Text = LangLookup(modLocalization.txslLangIndex.gnTestplanName)
        frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnVariant)
        frmMain.staDescriptions.Panels(frmMain.iCURRENTVARIANTNAME).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnCurrentTestplanVariant)
        frmMain.staDescriptions.Panels(frmMain.iCURRENTTESTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnTest)
        frmMain.staDescriptions.Panels(frmMain.iCURRENTTESTNAME).ToolTipText = LangLookup(modLocalization.txslLangIndex.gnCurrentTestName)
        frmMain.staDescriptions.Panels(frmMain.iCURRENTTESTNAME).Text = LangLookup(modLocalization.txslLangIndex.gnTest)
		frmMain.fraReport.Text = LangLookup(modLocalization.txslLangIndex.gnReport)
		frmMain.fraOperatorMessage.Text = LangLookup(modLocalization.txslLangIndex.gnOperatorMessages)
		frmMain.fraTestplanProgress.Text = LangLookup(modLocalization.txslLangIndex.gnTestplanProgress)
		frmMain.fraExecution.Text = LangLookup(modLocalization.txslLangIndex.gnTestplanExecution)
		frmMain.cmdRun.Text = LangLookup(modLocalization.txslLangIndex.gnRun1)
		frmMain.cmdStop.Text = LangLookup(modLocalization.txslLangIndex.gnStop1)
		frmMain.cmdAbort.Text = LangLookup(modLocalization.txslLangIndex.gnAbort1)
		frmMain.lblSerialNumber.Text = LangLookup(modLocalization.txslLangIndex.gnSerialNumber)
		frmMain.fraSystemStatus.Text = LangLookup(modLocalization.txslLangIndex.gnSystemStatus)
		frmMain.lblTitle.Text = LangLookup(modLocalization.txslLangIndex.gnTestExecSL)
		frmSelectVariant.Text = LangLookup(modLocalization.txslLangIndex.gnSelectVariant)
		frmSelectVariant.cmdCancel.Text = LangLookup(modLocalization.txslLangIndex.gnCancel)
		frmSelectVariant.cmdOK.Text = LangLookup(modLocalization.txslLangIndex.gnOK)
		frmSelectVariant.lblSelectANewVariant.Text = LangLookup(modLocalization.txslLangIndex.gnSelectanewVariant)
		frmSelectVariant.lblCurrentVariant.Text = LangLookup(modLocalization.txslLangIndex.gnCurrentVariant)
		
		frmSplash.Text = LangLookup(modLocalization.txslLangIndex.gnTestExecSL)
		frmSplash.lblCompany.Text = LangLookup(modLocalization.txslLangIndex.gnAgilentTechnologies)
		frmSplash.lblVersion.Text = LangLookup(modLocalization.txslLangIndex.gnVersion)
		
		frmSplash.lblProductName.Text = LangLookup(modLocalization.txslLangIndex.gnTestExecSL)
		
	End Sub
	
	
	Public Sub ChangeFonts()
		'Changes the fonts on the controls
		'This sub would most likely be called if you needed to change
		'to a font that would support non 8 bit languages, like Japanese.
		'This sub is positioned here (as opposed in modLocalization.bas), so that
		'modLocalization.bas can be application independent.
		'The code in this sub is very specific to the application.
		
		frmLoadTestplan.Font = VB6.FontChangeName(frmLoadTestplan.Font, "MS San Serif")
		frmLoadTestplan.cmdOK.Font = VB6.FontChangeName(frmLoadTestplan.cmdOK.Font, "MS San Serif")
		frmLoadTestplan.cmdCancel.Font = VB6.FontChangeName(frmLoadTestplan.cmdCancel.Font, "MS San Serif")
		frmLoadTestplan.lsvTestplanDetails.Font = VB6.FontChangeName(frmLoadTestplan.lsvTestplanDetails.Font, "MS San Serif")
		frmLoadTestplan.lblCurrentTestplan.Font = VB6.FontChangeName(frmLoadTestplan.lblCurrentTestplan.Font, "MS San Serif")
		frmLoadTestplan.lblTestplanDescription.Font = VB6.FontChangeName(frmLoadTestplan.lblTestplanDescription.Font, "MS San Serif")
		frmLoadTestplan.lblCurrentVariant.Font = VB6.FontChangeName(frmLoadTestplan.lblCurrentVariant.Font, "MS San Serif")
		frmLogin.Font = VB6.FontChangeName(frmLogin.Font, "MS San Serif")
		frmLogin.cmdCancel.Font = VB6.FontChangeName(frmLogin.cmdCancel.Font, "MS San Serif")
		frmLogin.cmdOK.Font = VB6.FontChangeName(frmLogin.cmdOK.Font, "MS San Serif")
		frmLogin.lblPassword.Font = VB6.FontChangeName(frmLogin.lblPassword.Font, "MS San Serif")
		frmLogin.lblName.Font = VB6.FontChangeName(frmLogin.lblName.Font, "MS San Serif")
		frmMain.Font = VB6.FontChangeName(frmMain.Font, "MS San Serif")
		frmMain.fraTestplanConfiguration.Font = VB6.FontChangeName(frmMain.fraTestplanConfiguration.Font, "MS San Serif")
		frmMain.chkReportPassedTests.Font = VB6.FontChangeName(frmMain.chkReportPassedTests.Font, "MS San Serif")
		frmMain.chkReportFailedTests.Font = VB6.FontChangeName(frmMain.chkReportFailedTests.Font, "MS San Serif")
		frmMain.chkReportException.Font = VB6.FontChangeName(frmMain.chkReportException.Font, "MS San Serif")
		frmMain.chkShowReport.Font = VB6.FontChangeName(frmMain.chkShowReport.Font, "MS San Serif")
		frmMain.staStats.Font = VB6.FontChangeName(frmMain.staStats.Font, "MS San Serif")
		frmMain.fraTxSLConfiguration.Font = VB6.FontChangeName(frmMain.fraTxSLConfiguration.Font, "MS San Serif")
		frmMain.cmdSelectVariant.Font = VB6.FontChangeName(frmMain.cmdSelectVariant.Font, "MS San Serif")
		frmMain.cmdLoadTestplan.Font = VB6.FontChangeName(frmMain.cmdLoadTestplan.Font, "MS San Serif")
		frmMain.cmdTxSLExit.Font = VB6.FontChangeName(frmMain.cmdTxSLExit.Font, "MS San Serif")
		frmMain.cmdLogin.Font = VB6.FontChangeName(frmMain.cmdLogin.Font, "MS San Serif")
		frmMain.lblBarCode.Font = VB6.FontChangeName(frmMain.lblBarCode.Font, "MS San Serif")
		frmMain.staDescriptions.Font = VB6.FontChangeName(frmMain.staDescriptions.Font, "MS San Serif")
		frmMain.fraReport.Font = VB6.FontChangeName(frmMain.fraReport.Font, "MS San Serif")
		frmMain.fraOperatorMessage.Font = VB6.FontChangeName(frmMain.fraOperatorMessage.Font, "MS San Serif")
		frmMain.fraTestplanProgress.Font = VB6.FontChangeName(frmMain.fraTestplanProgress.Font, "MS San Serif")
		frmMain.fraExecution.Font = VB6.FontChangeName(frmMain.fraExecution.Font, "MS San Serif")
		frmMain.cmdRun.Font = VB6.FontChangeName(frmMain.cmdRun.Font, "MS San Serif")
		frmMain.cmdStop.Font = VB6.FontChangeName(frmMain.cmdStop.Font, "MS San Serif")
		frmMain.cmdAbort.Font = VB6.FontChangeName(frmMain.cmdAbort.Font, "MS San Serif")
		frmMain.lblSerialNumber.Font = VB6.FontChangeName(frmMain.lblSerialNumber.Font, "MS San Serif")
		frmMain.fraSystemStatus.Font = VB6.FontChangeName(frmMain.fraSystemStatus.Font, "MS San Serif")
		frmMain.lblTitle.Font = VB6.FontChangeName(frmMain.lblTitle.Font, "MS San Serif")
		frmSelectVariant.Font = VB6.FontChangeName(frmSelectVariant.Font, "MS San Serif")
		'frmSelectVariant.lsvVariantDetails.ColumnHeaders("").Text
		frmSelectVariant.cmdCancel.Font = VB6.FontChangeName(frmSelectVariant.cmdCancel.Font, "MS San Serif")
		frmSelectVariant.cmdOK.Font = VB6.FontChangeName(frmSelectVariant.cmdOK.Font, "MS San Serif")
		frmSelectVariant.lblSelectANewVariant.Font = VB6.FontChangeName(frmSelectVariant.lblSelectANewVariant.Font, "MS San Serif")
		frmSelectVariant.lblCurrentVariant.Font = VB6.FontChangeName(frmSelectVariant.lblCurrentVariant.Font, "MS San Serif")
		
		frmSplash.Font = VB6.FontChangeName(frmSplash.Font, "MS San Serif")
		'frmSplash.lblCopyright.Font = "MS San Serif"
		frmSplash.lblCompany.Font = VB6.FontChangeName(frmSplash.lblCompany.Font, "MS San Serif")
		frmSplash.lblVersion.Font = VB6.FontChangeName(frmSplash.lblVersion.Font, "MS San Serif")
		'frmSplash.lblPlatform.Font = "MS San Serif"
		frmSplash.lblProductName.Font = VB6.FontChangeName(frmSplash.lblProductName.Font, "MS San Serif")
		'frmSplash.lblCompanyProduct.Font = "MS San Serif"
		
	End Sub
End Module