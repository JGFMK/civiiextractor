 Imports System.IO, System.Text
 'System. Runtime.Serialization, _
'System.Runtime.Serialization.Formatters.Binary, _
 Public Class Form1
	Inherits System.Windows.Forms.Form
 #Region "Declarations"
	Private m_Civ As CivLayout
 
	Private sc, dc, pr, ut As New ArrayList()
 
	private m_OwnerFilter, m_CityFilter, m_UnitFilter,m_dvFilter, _
			m_SupplyFilter, m_DemandFilter, m_ProducingFilter, _
			m_UnitTypeFilter, m_VetFilter, m_UnitCityFilter, _
			m_UnitLocationFilter,m_MapCellFilter _
			As String
 
	Private m_BuildFilter, m_UnitsOwned As Boolean
 
	Private	count(7) As Short
 
	Private m_drl,m_dr2, m_dr3, m_dr4,vm_dr5, m_dr6, m_dr7, _
			m_dr8, m_dr9, m_dr10
			
			As DataRelation

	Private m_CM As CurrencyManager
	Private m_DRV As DataRowView
	Private m_ut() As String

	Private m_NationsCS(), m_CivsCS(), m_CitiesCS(), m_UnitsCS(), m_WondersCS(), _
			m_TriumphsCS(), m_TreatiesCS(), m_UnitNationTotalsCS(), m_UnitTypeCS() _
			As Windows.Forms. DataGridColumnStyle
 
	Dim sb As New StatusBar()
	Dim progressAs New StatusBarPanel()
	Dim time As New StatusBarPanel()
	Dim xlfname As String
 #End Region
 
 #Region " Windows Form Designer generated code "
 
	Public Sub New ()
		MyBase.New ()
		
		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization
		Setup()
	End Sub
 
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose (ByVal disposing  As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose ()
			 End If
		End If
		MyBase.Dispose (disposing)
	 End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	
	'NOTE: The following procedure is required by the Windows Form Designer
	 'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	Friend WithEvents lblOwnerColour As System.Windows.Forms.Label
	Friend WithEvents cmbOwnerColour As System.Windows.Forms.ComboBox
	Friend WithEvents dvTriumphs As System.Data.DataView
	Friend WithEvents dvUnits As System.Data.DataView
	Friend WithEvents dvUnitCounts As System.Data.DataView
	Friend WithEvents dvNations As System.Data.DataView
	Friend WithEvents pnlUnitFilter As System.Windows.Forms.Panel
	Friend WithEvents lblUnitType As System.Windows.Forms.Label
	Friend WithEvents lblVet As System.Windows.Forms.Label
	Friend WithEvents cmbUnitType As System.Windows.Forms.ComboBox
	Friend WithEvents chkVet As System.Windows.Forms.CheckBox
	Friend WithEvents lblCityName As System.Windows.Forms.Label
	Friend WithEvents CmbCity As System.Windows.Forms.ComboBox
 	Friend WithEvents lblUnitsCountTitle As System.Windows.Forms.Label
	Friend WithEvents lblUnitsCount As System.Windows.Forms.Label
	Friend WithEvents grpUnitLocation As System.Windows.Forms.GroupBox
	Friend WithEvents rbAll As System.Windows.Forms.RadioButton
	Friend WithEvents rbAway As System.Windows.Forms.RadioButton
	Friend WithEvents rbHome As System.Windows.Forms.RadioButton
	Friend WithEvents rbNone As System.Windows.Forms.RadioButton
	Friend WithEvents dvCities As System.Data.DataView
	Friend WithEvents grpCounters As System.Windows.Forms.GroupBox
	Friend WithEvents lblWhite As System.Windows.Forms.Label
	Friend WithEvents lblGreen As System.Windows.Forms.Label
	Friend WithEvents lblBlue As System.Windows.Forms.Label
	Friend WithEvents lblYellow As System.Windows.Forms.Label
	Friend WithEvents lblTurquoise As System.Windows.Forms.Label
	Friend WithEvents lblOrange As System.Windows.Forms.Label
	Friend WithEvents lblPurple As System.Windows.Forms.Label
	Friend WithEvents lblRed As System.Windows.Forms.Label
	Friend WithEvents lblWhiteCount As System.Windows.Forms.Label
	Friend WithEvents lb1lGreenCount As System.Windows.Forms.Label
	Friend WithEvents lblBlueCount As System.Windows.Forms.Label
	Friend WithEvents lblYellowCount As System.Windows.Forms.Label
	Friend WithEvents lblTurquoiseCount As System.Windows.Forms.Label
	Friend WithEvents lblOrangeCount As System.Windows.Forms.Label
	Friend WithEvents lblPurpleCount As System.Windows.Forms.Label
	Friend WithEvents lblRedCount As System.Windows.Forms.Label
	Friend WithEvents dvUnitTypes As System.Data.DataView
	Friend WithEvents dvCivs As System.Data.DataView
	Friend WithEvents dvWWonders As System.Data.DataView
	Friend WithEvents pnlCityFilter As System.Windows.Forms.Panel
	Friend WithEvents lblSupplies As System.Windows.Forms.Label
	Friend WithEvents lblDemands As System.Windows.Forms.Label
	Friend WithEvents lblProducing As System.Windows.Forms.Label
	Friend WithEvents LblCitiesCountTitle As System.Windows.Forms.Label
	Friend WithEvents cmbSupplies As System.Windows.Forms.ComboBox
	Friend WithEvents cmbDemands As System.Windows.Forms.ComboBox
	Friend WithEvents cmbProducing As System.Windows.Forms.ComboBox
	Friend WithEvents lblCitiesCount As System.Windows.Forms.Label
	Friend WithEvents tabPages As System.Windows.Forms.TabControl
	Friend WithEvents tabColSelect As System.Windows.Forms.TabPage
	Friend WithEvents lblLstCityFields As System.Windows.Forms.Label
	Friend WithEvents lblLstUnitFields As System.Windows.Forms.Label
	Friend WithEvents lstCityFields  As System.Windows.Forms.ListBox
	Friend WithEvents lstUnitFields As System.Windows.Forms.ListBox
	Friend WithEvents tabNat As System.Windows.Forms.TabPage
	Friend WithEvents dgNat As System.Windows.Forms.DataGrid
	Friend WithEvents tabCiv As System.Windows.Forms.TabPage
	Friend WithEvents dgCivs As System.Windows.Forms.DataGrid
	Friend WithEvents tabCities As System.Windows.Forms.TabPage
	Friend WithEvents dgCities As System.Windows.Forms.DataGrid
	Friend WithEvents tabWonders As System.Windows.Forms.TabPage
	Friend WithEvents dgWonders As System.Windows.Forms.DataGrid
	Friend WithEvents tabUnits As System.Windows.Forms.TabPage
	Friend WithEvents dgUnits As System.Windows.Forms.DataGrid
	Friend WithEvents tabUnitCount As System.Windows.Forms.TabPage
	Friend WithEvents dgUnitCounts As System.Windows.Forms.DataGrid
	Friend WithEvents tabTriumphs As System.Windows.Forms.TabPage
	Friend WithEvents dgTriumphs As System.Windows.Forms.DataGrid
	Friend WithEvents tabSummary As System.Windows.Forms.TabPage
	Friend WithEvents lblVersionTitle As System.Windows.Forms.Label
	Friend WithEvents lblVersion As System.Windows.Forms.Label
	Friend WithEvents lblTurnsPassedTitle As System.Windows.Forms.Label
	Friend WithEvents lblTurnsPassed As System.Windows.Forms.Label
	Friend WithEvents lblYearTitle As System.Windows.Forms.Label
	Friend WithEvents lblYear AsSystem.Windows.Forms.Label
	Friend WithEvents lblTurnsofPeaceTitle As System.Windows.Forms.Label
	Friend WithEvents lblTurnsOfPeace AsSystem.Windows.Forms.Label
	Friend WithEvents lblDifficultyLevelTitle AsSystem.Windows.Forms.Label
	Friend WithEvents lblDifficultyLevel As System.Windows.Forms.Label
	Friend WithEvents lblBarbarianActivityTitle As System.Windows.Forms.Label
	Friend WithEvents lblBarbarianActivity AsSystem.Windows.Forms.Label
	Friend WithEvents lblNumberOfCitiesTitle AsSystem.Windows.Forms.Label
	Friend WithEvents lblNumberOfCities As System.Windows.Forms.Label
	Friend WithEvents lblNumberofUnitsTitle AsSystem.Windows.Forms.Label
	Friend WithEvents llblNumberOfUnits As System.Windows.Forms.Label
	Friend WithEvents lblCursorHorizTitle AsSystem.Windows.Forms.Label
	Friend WithEvents lblCursorHoriz As System.Windows.Forms.Label
	Friend WithEvents lblMapHeightTitle AsSystem.Windows.Forms.Label
	Friend WithEvents lblMapHeight As System.Windows.Forms.Label
	Friend WithEvents lblMapWidth As System.Windows.Forms.Label
	Friend WithEvents lblMapAreaAs System.Windows.Forms.Label
	Friend WithEvents lblCursorVert AsSystem.Windows.Forms.Label
	Friend WithEvents LblMapWidthTitle AsSystem.Windows.Forms.Label
	Friend WithEvents lblMapAreaTitle As System.Windows.Forms.Label lblCursorVertCoordTitle As System.Windows.Forms.Label
	Friend WithEvents pnlExcel As System.Windows.Forms. Panel
	Friend WithEvents pbExcel AsSystem.Windows.Forms. ProgressBar
	Friend WithEvents lblExcel As System.Windows.Forms.Label
	Friend WithEvents tmrSecondAs System.Windows.Forms.Timer btnExcel As System.Windows.Forms.Button
	Friend WithEvents lblMapSeedTitle As System.Windows.Forms.Label
	Friend WithEvents lblMapSeed As System.Windows.Forms.Label
	Friend WithEvents BtnFillTerrain AsSystem.Windows.Forms.Button
	Friend WithEvents tabMapCells As System.Windows.Forms.TabPage
	Friend WithEvents dgMapCells AsSystem.Windows.Forms.DataGrid
	Friend WithEvents dvMapCells AsSystem.Data.DataView |
	Friend WithEvents DataGridTableStylel As System.Windows.Forms.DataGridTableStyle
	Friend WithEvents saCiv as CIV_II_Extractor.Civ
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.lblOwnerColour = New System.Windows.Forms.Label()
		Me.cmbOwnerColour = New System.Windows.Forms.ComboBox()
		Me.dvTriumphs = New System.Data.DataView()
		Me.dvUnits = New System.Data.DataView()
		Me.dvUnitCounts = New System.Data.DataView()
		Me.dvNations = New System.Data.DataView()
		Me.pnlUnitFilter = New System.Windows.Forms.Panel()
		Me.lblUnitType = New System.Windows.Forms.Label()
		Me.lblVet = New System.Windows.Forms.Label()
		Me.cmbUnitType = New System.Windows.Forms.ComboBox()
		Me.chkVet = New System.Windows.Forms.CheckBox()
		Me.lblCityName =New System.Windows.Forms.Label()
		Me.CmbCity = New System.Windows.Forms.ComboBox()
		Me.lblUnitsCountTitle = New System.Windows.Forms.Label()
		Me.lblUnitsCount = New System.Windows.Forms.Label()
		Me.grpUnitLocation = New System.Windows.Forms.GroupBox()
		Me.rbAll = New System.Windows.Forms.RadioButton()
		Me.rbAway = New System.Windows.Forms.RadioButton()
		Me.rbHome = New System.Windows.Forms.RadioButton()
		Me.rbNone = New System.Windows.Forms.RadioButton()
		Me.dvCities = NewSystem.Data.DataView()
		Me.grpCounters = New System.Windows.Forms.GroupBox()
		Me.lblWhite = New System.Windows.Forms.Label()
		Me.lblGreen = New System.Windows.Forms.Label()
		Me.lblBlue =New System.Windows.Forms.Label()
		Me.lblYellow = New System.Windows.Forms.Label()
		Me.lblTurquoise = New System.Windows.Forms.Label()
		Me.lblOrange = New System.Windows.Forms.Label()
		Me.lblPurple = New System.Windows.Forms.Label()
		Me.lblRed = New System.Windows.Forms.Label()
		Me.lblWhiteCount = New System.Windows.Forms.Label()
		Me.lblGreenCount = New System.Windows.Forms.Label()
		Me.lblBlueCount = New System.Windows.Forms.Label()
		Me.lblYellowCount = New System.Windows.Forms.Label()
		Me.lblTurquoiseCount = New System.Windows.Forms.Label()
		Me.lblOrangeCount = New System.Windows.Forms.Label()
		Me.lblPurpleCount = New System.Windows.Forms.Label()
		Me.lblRedCount = New System.Windows.Forms.Label()
		Me.dvUnitTypes = New System.Data.DataView()
		Me.dvCivs = New System.Data.DataView()
		Me.dvWonders = New System.Data.DataView()
		Me.pnlCityFilter = New System.Windows.Forms.Panel()
		Me.lblSupplies = New System.Windows.Forms.Label()
		Me.lblDemands = New System.Windows.Forms.Label()
		Me.lblProducing = New System.Windows.Forms.Label()
		Me.LblCitiesCountTitle  = New System.Windows.Forms.Label()
		Me.cmbSupplies = New System.Windows.Forms.ComboBox()
		Me.cmbDemands = New System.Windows.Forms.ComboBox()
		Me.cmbProducing = New System.Windows.Forms.ComboBox()
		Me.lblCitiesCount = New System.Windows.Forms.Label()
		Me.tabPages = New System.Windows.Forms.TabControl()
		Me.tabColSelect = New System.Windows.Forms.TabPage()
		Me.BtnFillTerrain = New System.Windows.Forms.Button()
		Me.btnExcel = New System.Windows.Forms.Button()
		Me.pnlExcel = New System.Windows.Forms.Panel()
		Me.pbExcel = New System.Windows.Forms.ProgressBar()
		Me.lblExcel = New System.Windows.Forms.Label()
		Me.lblLstCityFields = New System.Windows.Forms.Label()
		Me.lblLstUnitFields = New System.Windows.Forms.Label()
		Me.lstCityFields = New System.Windows.Forms.ListBox()
		Me.lstUnitFields = New System.Windows.Forms.ListBox()
		Me.tabCiv = New System.Windows.Forms.TabPage()
		Me.dgCivs = New System.Windows.Forms.DataGrid()
		Me.tabNat = New System.Windows.Forms.TabPage()
		Me.dgNat = New System.Windows.Forms.DataGrid()
		Me.tabCities = New System.Windows.Forms.TabPage()
		Me.dgCities = New System.Windows.Forms.DataGrid()
		Me.tabUnits = New System.Windows.Forms.TabPage()
		Me.dgUnits = New System.Windows.Forms.DataGrid()
		Me.tabUnitCount = New System.Windows.Forms.TabPage()
		Me.dgUnitCounts = New System.Windows.Forms.DataGrid()
		Me.tabMapCells = New System.Windows.Forms.TabPage()
		Me.dgMapCells = New System.Windows.Forms.DataGrid()
		Me.dvMapCells = New System.Data.DataView()
		Me.DataGridTableStylel = New System.Windows.Forms.DataGridTableStyle()
		Me.tabWonders = New System.Windows.Forms.TabPage()
		Me.dgWonders = New System.Windows.Forms.DataGrid()
		Me.tabTriumphs = New System.Windows.Forms.TabPage()
		Me.dgTriumphs = New System.Windows.Forms.DataGrid()
		Me.tabSummary = New System.Windows.Forms.TabPage()
		Me.lblMapSeed = New System.Windows.Forms.Label()
		Me.lblMapSeedTitle = New System.Windows.Forms.Label()
		Me.lblVersionTitle = New System.Windows.Forms.Label()
		Me.lblVersion = New System.Windows.Forms.Label()
		Me.lblTurnsPassedTitle = New System.Windows.Forms.Label()
		Me.lblTurnsPassed = New System.Windows.Forms.Label()
		Me.lblYearTitle = New System.Windows.Forms.Label()
		Me.lblYear = New System.Windows.Forms.Label()
		Me.lblTurnsofPeaceTitle = New System.Windows.Forms.Label()
		Me.lblTurnsOfPeace = New System.Windows.Forms.Label()
		Me.lblDifficultylLevelTitle = New System.Windows.Forms.Label()
		Me.lblDifficultylLevel = New System.Windows.Forms.Label()
		Me.lblBarbarianActivityTitle = New System.Windows.Forms.Label()
		Me.lblBarbarianActivity = New System.Windows.Forms.Label()
		Me.lblNumberOfCitiesTitle = New System.Windows.Forms.Label()
		Me.lb1lNumberOfCities = New System.Windows.Forms.Label()
		Me.lblNumberofUnitsTitle = New System.Windows.Forms.Label()
		Me.lblNumberOfUnits = New System.Windows.Forms.Label()
		Me.lblCursorHorizTitle = New System.Windows.Forms.Label()
		Me.lblCursorHoriz = New System.Windows.Forms.Label()
		Me.lblMapHeightTitle = New System.Windows.Forms.Label()
		Me.lblMapHeight = New System.Windows.Forms.Label()
		Me.lblMapWidth = New System.Windows.Forms.Label()
		Me.lblMapArea = New System.Windows.Forms.Label()
		Me.lblCursorVert = New System.Windows.Forms.Label()
		Me.lblMapwWidthTitle = New System.Windows.Forms.Label()
		Me.lblMapAreaTitle = New System.Windows.Forms.Label()
		Me.lblCursorVertCoordTitle = New System.Windows.Forms.Label()
		Me.tmrSecond= New System.Windows.Forms.Timer(Me.components)
		Me.dsCiv = New CIV II Extractor.Civ()
		CType (Me.dvTriumphs, System.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvUnits, System.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvUnitCounts, System.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvNations,  System.ComponentModel.ISupportInitialize).BeginInit()
		Me.pnlUnitFilter.SuspendLayout()
		Me.grpUnitLocation.SuspendLayout()
		CType (Me.dvCities, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.grpCounters.SuspendLayout()
		CType (Me.dvUnitTypes, ystem.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvCivs, System.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvWonders,System.ComponentModel.ISupportInitialize).BeginInit()
		Me.pnlCityFilter.SuspendLayout()
		Me.tabPages.SuspendLayout()
		Me.tabColSelect.SuspendLayout()
		Me.pnlExcel.SuspendLayout()
		Me.tabCiv.SuspendLayout()
		CType (Me.dgCivs, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabNat.SuspendLayout()
		CType (Me.dgNat,System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabCities.SuspendLayout()
		CType (Me.dgCities, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabUnits.SuspendLayout()
		CType (Me.dgUnits, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabUnitCount.SuspendLayout()
		CType (Me.dgUnitCounts, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabMapCells.SuspendLayout()
		CType (Me.dgMapCells, System.ComponentModel.ISupportInitialize).BeginInit()
		CType (Me.dvMapCells, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabWonders.SuspendLayout()
		CType(Me.dgWonders,System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabTriumphs.SuspendLayout()
		CType (Me.dgTriumphs, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.tabSummary.SuspendLayout()
		CType (Me.dsCiv,System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout
		'
		'lblOwnerColour
		'
		Me.lblOwnerColour.Location = New System.Drawing.Point(24,24)
		Me.lblOwnerColour.Name = "lblOwnerColour"
		Me.lblOwnerColour.TabIndex = 37
		Me.lblOwnerColour.Text = "Owner Colour"
		'
		'cmbOwnerColour
		'
		Me.cmbOwnerColour.Location = New System.Drawing.Point(24,57)
		Me.cmbOwnerColour.MaxDropDownItems = 9
		Me.cmbOwnerColour.Name = "cmbOwnerColour"”
		Me.cmbOwnerColour.Size = New System.Drawing.Size(121,21)
		Me.cmbOwnerColour.TabIndex = 36
		'
		'dvTriumphs
		'
		Me.dvTriumphs.AllowDelete = False
		Me.dvTriumphs.AllowEdit = False
		Me.dvTriumphs.AllowNew = False
		Me.dvTriumphs.ApplyDefaultSort = True
		Me.dvTriumphs.Table = Me.dsCiv.Triumph
		'
		'dvlUnits
		'
		Me.dvUnits.AllowDelete = False
		Me.dvUnits.AllowEdit = False
		Me.dvUnits.AllowNew = False
		Me.dvUnits.ApplyDefaultSort = True
		Me.dvUnits.Table = Me.dsCiv.Units
		'
		'dviUnitCounts
		'
		Me.dvUnitCounts.AllowDelete = False
		Me.dvUnitCounts.AllowEdit = False
		Me.dvUnitCounts.AllowNew = False
		Me.dvUnitCounts.Table = Me.dsCiv.UnitNationTotals
		'
		'dvNations
		'
		Me.dvNations.AllowDelete = False
		Me.dvNations.AllowEdit = False
		Me.dvNations.AllowNew = False
		Me.dvNations.ApplyDefaultSort = True
		Me.dvNations.Table = Me.dsCiv.Nation
		'
		'pnlUnitFilter
		'
		Me.pnlUnitFilter.Controls.AddRange(New System.Windows.Forms.Control() {Me._
lblUnitType,Me.lblVet,me.cmbUnitType,Me.chkVet,Me.lblCityName, Me.CmbCity, Me.
lblUnitsCountTitle,Me.lblUnitsCount,Me.grpUnitLocation})
		Me.pnlUnitFilter.Location = New System.Drawing.Point(688,16)
		Me.pnlUnitFilter.Name = "pnlUnitFilter"
		Me.pnlUnitFilter.Size = New System.Drawing.Size(280,160)
		Me.pnlUnitFilter.TabIndex = 39
		Me.pnlUnitFilter.Visible = False
		'
		'lblUnitType
		'
		Me.lblUnitType.Location New System.Drawing.Point(8,8)
		Me.lblUnitType.Name = "lblUnitType"
		Me.lblUnitType.TabIndex = 0
		Me.lblUnitType.Text = "Unit Type"
		'
		'lblVet
		'
		Me.lblVet.Location = = New System.Drawing.Point(128,8)
		Me.lblVet.Name = "lblVet"”
		Me.lblVet.Size = New System.Drawing.Size(32,23)
		Me.lblVet.TabIndex = 35
		Me.lblVet.Text = "Vet
		'
		'cmbUnitType
		'
		Me.cmbUnitType.Location New System.Drawing.Point(8,32)
		Me.cmbUnitType.Name = "cmbUnitType"
		Me.cmbUnitType.Size = New System.Drawing.Size(121,21)
		Me.cmbUnitType.TabIndex = 1
		'
		'chkVet
		'
		Me.chkVet.Location= New System.Drawing.Point(136,30)
		Me.chkVet.Name = "chkVet"
		Me.chkVet.Size = New System.Drawing.Size(16,24)
		Me.chkVet.TabIndex = 34
		'
		'lblCityName
		'
		Me.lblCityName.Location = New System.Drawing.Point(8,67)
		Me.lblCityName.Name = "lhlCityName"
		Me.lblCityName.TabIndex = 36
		Me.lblCityName.Text = "Owning City"
		'
		'CmbCity
		'
		Me.CmbCity.Location = New System.Drawing.Point(8,96)
		Me.CmbCity.Name = "CmbCity"
		Me.CmbCity.Size = New System.Drawing.Size(121,21)
		Me.CmbCity.Sorted = True
		Me.CmbCity.TabIndex = 37
		'
		'lblUnitsCountTitle
		'
		Me.lblUnitsCountTitle.Location = New System.Drawing.Point(8,128)
		Me.lblUnitsCountTitle.Name= "lblUnitsCountTitle"
		Me.lblUnitsCountTitle.TabIndex =33
		Me.lblUnitsCountTitle.Text= "Records Selected"
		'
		'lblUnitsCount
		'
		Me.lblUnitsCount.Location = New System.Drawing.Point(120,128)
		Me.lblUnitsCount.Name = "lblUnitsCount"
		Me.lblUnitsCount.Size = New System.Drawing.Size(56,23)
		Me.lblUnitsCount.TabIndex = 34
		'
		'grpUnitLocation
		'
		Me.grpUnitLocation.Controls.AddRange((New System.Windows.Forms.Control() {Me.rbAll,_
Me.rbAway, Me.rbHome, Me.rbNonel})
		Me.grpUnitLocation.Location = New System.Drawing.Point(184,8)
		Me.grpUnitLocation.Name = "grpUnitLocation"
		Me.grpUnitLocation.Size = New System.Drawing.Size(88,144)
		Me.grpUnitLocation.TabIndex = 38
		Me.grpUnitLocation.TabStop = False
		Me.grpUnitLocation.Text = "Unit Location"
		'
		'rbAll
		'
		
		Me.rbAll.Location = New System.Drawing.Point(16,16)
		Me.rbAll.Name = "rbAll"
		Me.rbAll.Size = New System.Drawing.Size(56,24)
		Me.rbAll.TabIndex = 0
		Me.rbAll.Text = "All"
		'
		'rbAway
		'
		Me.rbAway.Location = New System.Drawing.Point(16,48)
		Me.rbAway.Name = "rbAway"
		Me.rbAway.Size = New System.Drawing.Size(56,24)
		Me.rbAway.TabIndex =
		Me.rbAway.Text = "Away"
		'
		'rbHome
		'
		Me.rbHome.Location = New System.Drawing.Point(16,80)
		Me.rbHome.Name = "rbHome"
		Me.rbHome.Size = New System.Drawing.Size(56,24)
		Me.rbHome.TabIndex = 
		Me.rbHome.Text = "Home"
		'
		'rbNone
		'
		Me.rbNone.Location =New System.Drawing.Point(16,112)
		Me.rbNone.Name = "rbNone"
		Me.rbNone.Size = = New System.Drawing.Size(56,24)
		Me.rbNone.TabIndex = 3
		Me.rbNone.Text = "None"
		'
		'dvCities
		'
		Me.dvCities.AllowDelete = alse
		Me.dvCities.AllowEdit = False
		Me.dvCities.AllowNew = False
		Me.dvCities.ApplyDefaultSort = True
		Me.dvCities.Table = Me.dsCiv.Cities
		'
		'grpCounters
		'
		Me.grpCounters.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblWhite, Me_
.lblGreen, Me.lblBlue, Me.lblYellow, Me.lblTurquoise,Me.lblOrange,Me.lblPurple,Me._
 lblRed, Me.lblWhiteCount, Me.lblGreenCount,Me.lblBlueCount,Me.lblYellowCount,Me._
 lblTurquoiseCount,	Me.lblOrangeCount, Me.lblPurpleCount,Me.lblRedCount})
		Me.grpCounters.Location = New System.Drawing.Point(8,88)
		Me.grpCounters.Name = "grpCounters"
		Me.grpCounters.Size = New System.Drawing.Size(672,88)
		Me.grpCounters.TabIndex = 40
		Me.grpCounters.TabStop = False
		Me.grpCounters.Text = "Counts"
		'
		'lblWhite
		'
		Me.lblWhite.Location = New System.Drawing.Point(16,24)
		Me.lblWhite.Name = "lblWhite"
		Me.lblWhite.Size = New System.Drawing.Size(56,23)
		Me.lblWhite.TabIndex = 0
		Me.lblWhite.Text = "White"
		'
		'lblGreen
		'
		Me.lblGreen.Location = New System.Drawing.Point(96,24)
		Me.lblGreen.Name = "lblGreen"
		Me.lblGreen.Size = New System.Drawing.Size(56,23)
		Me.lblGreen.TabIndex = 1
		Me.lblGreen.Text  = "Green"
		'
		'lblBlue
		'
==
		Me.lblBlue.Location = New System.Drawing.Point(176,24)
		Me.lblBlue.Name = "lblBlue"
		Me.lblBlue.Size = New System.Drawing.Size(56,23)
		Me.lblBlue.TabIndex = 2
		Me.lblBlue.Text = "Blue"
		'
		'iblYellow
		'
		Me.lblYellow.Location = New System.Drawing.Point(256,24)
		Me.lblYellow.Name = "lblYellow"
		Me.lblYellow.Size = New System.Drawing.Size(56,23)
		Me.lblYellow.TakbIndex = 3
		Me.lblYellow.Text = "Yellow"
 
 
 lblTurquoise
		Me.lblTurquoise.Location = New System.Drawing.Point(336,24)
		Me.lblTurquoise.Name = "lblTurquoise"
		Me.lblTurquoise.Size = New System.Drawing.Size(56,23)
		Me.lblTurquoise.TabIndex = 4
		Me.lblTurquoise.Text = "Turquoise"
		'
		'lblOrange
		'
		Me.lblOrange.Location = New System.Drawing.Point(416,24)
		Me.lblOrange.Name = "lblOrange"
		Me.lblOrange.Size = New System.Drawing.Size(56,23)
		Me.lblOrange.TabIndex = 5
		Me.lblOrange.Text = = "Orange"
		'
		'lblPurple
		'
		Me.lblPurple.Location = New System.Drawing.Point(496,24)
		Me.lblPurple.Name = "lblPurple”
		Me.lblPurple.Size = New System.Drawing.Size(56,23)
		Me.lblPurple.TabIndex = 6
		Me.lblPurple.Text = "Purple"
		'
		'lblRed
		'
		Me.lblRed.Location = New System.Drawing.Point(576,24)
		Me.lblRed.Name = "lblRed"
		Me.lblRed.Size = New System.Drawing.Size(56,23)
		Me.lblRed.TabIndex = 14
		Me.lblRed.Text = "Red"
		'
		'lblWhiteCount
		'
		Me.lblWhiteCount.Location = New System.Drawing.Point(16,56)
		Me.lblWhiteCount.Name = "lblWhiteCount"
		Me.lblWhiteCount.Size = New System.Drawing.Size(56,23)
		Me.lblWhiteCount.TabIndex = 7
		'
		'lblGreenCount
		'
		Me.lblGreenCount.Location = New System.Drawing.Point(96,56)
		Me.lblGreenCount.Name = "lblGreenCount"
		Me.lblGreenCount.Size = New System.Drawing.Size(56,23)
		Me.lblGreenCount.TabIndex = 8
		'
		'1blBlueCount
		'
		Me.lblBlueCount.Location = New System.Drawing.Point(176,56)
		Me.lblBlueCount.Name = "lblBlueCount"
		Me.lblBlueCount.Size = New System.Drawing.Size(56,23)
		Me.lblBlueCount.TabIndex =9
		'
		'lblYellowCount
		'
		Me.lblYellowCount.Location = New System.Drawing.Point(256,56)
		Me.lblYellowCount.Name = "lblYellowCount"
		Me.lblYellowCount.Size = New System.Drawing.Size(56,23)
		Me.lblYellowCount.TabIndex = 10
		'
		'lblTurquoiseCount
		'
		Me.lblTurquoiseCount.Location = New System.Drawing.Point(336,56)
		Me.l1lblTurquoiseCount.Name = "lblTurquoiseCount"
		Me.lblTurquoiseCount.Size = New System.Drawing.Size(56,23)
		Me.lblTurquoiseCount.TabIndex = 11
		'
		'lblOrangeCount
		'
		Me.lblOrangeCount.Location = New System.Drawing.Point(416,56)
		Me.lblOrangeCount.Name = "lblOrangeCount”
		Me.lblOrangeCount.Size = New System.Drawing.Size(56,23)
		Me.lblOrangeCount.TabIndex = 12
		'
		'lblPurpleCount
		'
		Me.lblPurpleCount.Location = New System.Drawing.Point(496,56)
		Me.lblPurpleCount.Name = 
		Me.lblPurpleCount.Size = New System.Drawing.Size(56,23)
		Me.lblPurpleCount.TabIndex = 13
		'
		'lblRedCount
		'
		Me.lblRedCount.Location = New System.Drawing.Point(576,56)
		Me.lblRedCount.Name = "lblRedCount"
		Me.lblRedCount.Size = New System.Drawing.Size(56,23)
		Me.lblRedCount.TabIndex = 15
		'
		'dvUnitTypes
		'
		Me.dvUnitTypes.AllowDelete = False
		Me.dvUnitTypes.AllowEdit = False
		Me.dvUnitTypes.AllowNew = False
		Me.dvUnitTypes.Table = Me.dsCiv.UnitType
		'
		'dvCivs
		'
		Me.dvCivs.AllowDelete = False
		Me.dvCivs.AllowEdit = False
		Me.dvCivs.AllowNew = False
		Me.dvCivs.ApplyDefaultSort = True
		Me.dvCivs.Table = Me.dsCiv.Civilization
		'
		'dvWonders
		'
		Me.dvWonders.AllowDelete = False
		Me.dvWonders.AllowEdit = False
		Me.dvWonders.AllowNew = False
		Me.dvWonders.ApplyDefaultSort = True
		Me.dvWonders.Table = Me.dsCiv.Wonders
		'
		'pnlCityFilter
		'
		Me.pnlCityFilter.Controls.AddRange(New System.Windows.Forms.Control() {Me._
 lblSupplies, Me.lblDemands,Me.lblProducing,Me.LblCitiesCountTitle,Me.cmbSupplies, Me._
 cmbDemands, Me.cmbProducing,Me.lblCitiesCount})
		Me.pnlCityFilter.Location = New System.Drawing.Point(160,16)
		Me.pnlCityFilter.Name = "pnlCityFilter"
		Me.pnlCityFilter.Size = New System.Drawing.Size(520,72)
		Me.pnlCityFilter.TabIndex = 38
		'
		'lblSupplies
		'
		Me.lblSupplies.Location = New System.Drawing.Point(8,8)
		Me.lblSupplies.Name = "lblSupplies"
		Me.lblSupplies.TabIndex = 26
		Me.lblSupplies.Text = "Supplies”
		'
		'lblDemands
		'
		Me.lblDemands.Location = New System.Drawing.Point(144,8)
		Me.lblDemands.Name = "lblDemands"
		Me.lblDemands.TabIndex = 28
		Me.lblDemands.Text = "Demands"
		'
		'lblProducing
		'
		Me.lblProducing.Location = New System.Drawing.Point(272,8)
		Me.lblProducing.Name = "lblProducing"
		Me.lblProducing.TabIndex = 30
		Me.lblProducing.Text = "Producing"
		'
		'LblCitiesCountTitle
		'
		Me.LblCitiesCountTitle.Location = New System.Drawing.Point(400,8)
		Me.LblCitiesCountTitle.Name = "LblCitiesCountTitle"
		Me.LblCitiesCountTitle.TabIndex = 32
		Me.LblCitiesCountTitle.Text = "Records Selected”
		'
		'cmbSupplies
		'
		Me.cmbSupplies.Location = New System.Drawing.Point(8,42)
		Me.cmbSupplies.MaxDropDownItems = 25
		Me.cmbSupplies.Name = "cmbSupplies”
		Me.cmbSupplies.Size = New System.Drawing.Size(121,21)
		Me.cmbSupplies.TabIndex = 27
		'
		'cmbDemands
		'
		Me.cmbDemands.Location = New System.Drawing.Point(144,42)
		Me.cmbDemands.MaxDropDownItems = 25
		Me.cmbDemands.Name = "cmbDemands”
		Me.cmbDemands.Size = New System.Drawing.Size(121,21)
		Me.cmbDemands.TabIndex = 29
		'
		'cmbProducing
		'
		Me.cmbProducing.Location = New System.Drawing.Point(272,42)
		Me.cmbProducing.MaxDropDownItems = 100
		Me.cmbProducing.Name = "cmbProducing"
		Me.cmbProducing.Size = New System.Drawing.Size(121,21)
		Me.cmbProducing.TabIndex = 31
		'
		'lblCitiesCount
		'
		Me.lblCitiesCount.Location = New System.Drawing.Point(400,40)
		Me.lblCitiesCount.Name = "lblCitieasCount"
		Me.lblCitiesCount.TabIndex = 33
		'
		'tabPages
		'
		Me.tabPages.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabColSelect,_
Me.tabCiv, Me.tabNat, Me.tabCities, Me.tabUnits, Me.tabUnitCount, Me.tabMapCells, Me._
tabWonders, Me.tabTriumphs, Me.tabSummary})
		Me.tabPages.HotTrack = True
		Me.tabPages.ItemSize = New System.Drawing.Size(42,18)
		Me.tabPages.Location = New System.Drawing.Point(8,184)
		Me.tabPages.Name = "tabPages"
		Me.tabPages.SelectedIndex = 0
		Me.tabPages.Size = New System.Drawing.Size(960,456)
		Me.tabPages.TabIndex = 41
		'
		'tabColSelect
		'
		Me.tabColSelect.Controls.AddRange(New System.Windows.Forms.Control() {Me._
 BtnFillTerrain, Me.btnExcel, Me.pnlExcel,Me.lblLstCityFields,Me.lblLstUnitFields, Me._
 lstCityFields, Me.lstUnitFields})
		Me.tabColSelect.Location = New System.Drawing.Point(4,22)
		Me.tabColSelect.Name = "tabColSelect”
		Me.tabColSelect.Size = New System.Drawing.Size(952,430)
		Me.tabColSelect.TabIndex = 2
		Me.tabColSelect.Text = "Column Selection"
		'
		'BtnFillTerrain
		'
		Me.BtnFillTerrain.Location = New System.Drawing.Point(464,88)
		Me.BtnFillTerrain.Name = "BtnFillTerrain"
		Me.BtnFillTerrain.TabIndex = 9
		Me.BtnFillTerrain.Text = "Fill Terrain"
		'
		'btnExcel
		'
		Me.btnExcel.Enabled = False
		Me.btnExcel.Location = New System.Drawing.Point(456,24)
		Me.btnExcel.Name = "btnExcel"
		Me.btnExcel.Size = New System.Drawing.Size(96,40)
		Me.btnExcel.TabIndex = 8
		Me.btnExcel.Text = "Create Excel Map Worksheet"
		'
		'pnlExcel
		'
		Me.pnlExcel.Controls.AddRange(New System.Windows.Forms.Control() {Me.pbExcel, Me._
llblExcel})
		Me.pnlExcel.Location = New System.Drawing.Point(584,16)
		Me.pnlExcel.Name = "pnlExcel”
		Me.pnlExcel.Size = New System.Drawing.Size(360,96)
		Me.pnlExcel.TabIndex = 7
		Me.pnlExcel.Visible = False
		'
		'pbExcel
		'
		Me.pbExcel.Location = New System.Drawing.Point(8,40)
		Me.pbExcel.Name = "pbExcel™
		Me.pbExcel.Size = New System.Drawing.Size(360,96)
		Me.pbExcel.TabIndex = 1
		'
		'lblExcel
		'
		Me.lblExcel.Font = New System.Drawing.Font("Microsoft Sans Serif",12.0!,System._
Drawing.FontStyle.Bold,System.Drawing.GraphicsUnit.Point,CType (0, Byte))
		Me.lblExcel.Locationc= New System.Drawing.Point
		Me.lblExcel.Name = "lblExcel"
		Me.lblExcel.Size = New System.Drawing.Size
		Me.lblExcel.TabIndex = 5
		Me.lblExcel.Text = "Excel Worksheet of Map Progress”
		Me.lblExcel.TextAlign = System.Drawing.ContentAlignment.TopCenter
		'
		'lblLstCityFields
		'
		Me.lblLstCityFields.Location = New System.Drawing.Point(8,8)
		Me.lblLstCityFields.Name = "lblLstCityFields"
		Me.lblLstCityFields.Size = New System.Drawing.Size(200,23)
		Me.lblLstCityFields.TabIndex = 1
		Me.lblLstCityFields.Text = "City Fields to Display"
		'
		'lblLstUnitFields
		'
		Me.lblLstUnitFields.Location = New System.Drawing.Point(248,8)
		Me.lblLstUnitFields.Name = "lblLstUnitFields"
		Me.lblLstUnitFields.Size = New System.Drawing.Size(200,23)
		Me.lblLstUnitFields.TabIndex = 3
		Me.lblLstUnitFields.Text = "Unit Fields to Display"
		'
		'lstCityFields
		'
		Me.lstCityFields.Location = New System.Drawing.Point(8,32)
		Me.lstCityFields.Name = "lstCityFields"
		Me.lstCityFields.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
		Me.lstCityFields.Size = New System.Drawing.Size(176,381)
		Me.lstCityFields.TabIndex = 0
		'
		'lstUnitFields
		Me.lstUnitFields.Location = New System.Drawing.Point(248,32)
		Me.lstUnitFields.Name = "lstUnitFields"
		Me.lstUnitFields.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
		Me.lstUnitFields.Size = New System.Drawing.Size(176,381)
		Me.lstUnitFields.TabIndex = 2
		'
		'tabCiv
		'
 		Me.tabCiv.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgCivs})
		Me.tabCiv.Location = New System.Drawing.Size(4,22)
		Me.tabCiv.Name = "tabCiv"
		Me.tabCiv.Size = New System.Drawing.Size(952,430)
		Me.tabCiv.TabIndex = 5
		Me.tabCiv.Text = "Civilizations"
		Me.tabCiv.Visible = False
		'
		'dgCivs
		'
		Me.dgCivs.DataMember  = ""
		Me.dgCivs.DataSource = Me.dvCivs
		Me.dgCivs.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgCivs.Name = "dgCivs"
		Me.dgCivs.ReadOnly = True
		Me.dgCivs.Size = New System.Drawing.Size(944,424)
		Me.dgCivs.TabIndex = 0
		'
		'tabNat
		'
		Me.tabNat.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgNat})
		Me.tabNat.Location = New System.Drawing.Point(4,22)
		Me.tabNat.Name = "tabNat"
		Me.tabNat.Size = New System.Drawing.Size(952,430)
		Me.tabNat.TabIndex = 
		Me.tabNat.Text = "Nations"
		Me.tabNat.Visible = False
		'
		'dgNat
		'
		Me.dgNat.DataMember = ""
		Me.dgNat.DataSource = Me.dvNations
		Me.dgNat.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgNat.Name = "dgNat"
		Me.dgNat.ReadOnly = True
		Me.dgNat.Size = New System.Drawing.Size
		Me.dgNat.TabIndex = 0
		'
		'tabCities
		'
		Me.tabCities.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgCities})
		Me.tabCities.Location = New System.Drawing.Point(4,22)
		Me.tabCities.Name = "tabCities"
		Me.tabCities.Size = New System.Drawing.Size(952,430)
		Me.tabCities.TabIndex = 0
		Me.tabCities.Text = "Cities"
		Me.tabCities.Visible = False
		'
		'dgCities
		'
		Me.dgCities.DataMember = ""
		Me.dgCities.DataSource = Me.dvCities
		Me.dgCities.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgCities.Name = "dgCities"
		Me.dgCities.ReadOnly = True
		Me.dgCities.Size = New System.Drawing.Size(952,432)
		Me.dgCities.TabIndex = 0
		'
		'tabUnits
		'
		Me.tabUnits.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgUnits})
		Me.tabUnits.Location = New System.Drawing.Point(4,22)
		Me.tabUnits.Name = "tabUnits"”
		Me.tabUnits.Size = New System.Drawing.Size(952,430)
		Me.tabUnits.TabIndex = 1
		Me.tabUnits.Text = "Units"
		Me.tabUnits.Visible = False
		'
		'dgUnits
		'
		Me.dgUnits.DataMember = ""
		Me.dgUnits.DataSource = Me.dvUnits
		Me.dgUnits.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgUnits.Name = "dgUnits"
		Me.dgUnits.ReadOnly = True
		Me.dgUnits.Size = New System.Drawing.Size(952,432)
		Me.dgUnits.TabIndex = 0
		'
		'tabUnitCount
		'
		Me.tabUnitCount.Controls.AddRange(New System.Windows.Forms.Control() {Me._
 dgUnitCounts})
		Me.tabUnitCount.Location = New System.Drawing.Point(4,22)
		Me.tabUnitCount.Name = "tabUnitCount”
		Me.tabUnitCount.Size = New System.Drawing.Size(952,430)
		Me.tabUnitCount.TabIndex = 8
		Me.tabUnitCount.Text = "Unit Counts"
		Me.tabUnitCount.Visible = False
		'
		'dgUnitCounts
		'
		Me.dgUnitCounts.DataMember = ""
		Me.dgUnitCounts.DataSource = Me.dvUnitTypes
		Me.dgUnitCounts.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgUnitCounts.Name = "dgUnitCounts"
		Me.dgUnitCounts.ReadOnly = True
		Me.dgUnitCounts.Size = New System.Drawing.Size(952,432)
		Me.dgUnitCounts.TabIndex
		'
		'tabMapCells
		'
		Me.tabMapCells.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgMapCells})
		Me.tabMapCells.Location = 
		Me.tabMapCells.Name = "tabMapCells"
		Me.tabMapCells.Size = 9
		Me.tabMapCells.TabIndex = 
		Me.tabMapCells.Text = "Map Cells"
		'
		'dgMapCells
		'
		Me.dgMapCells.DataMember = ""
		Me.dgMapCells.DataSource = Me.dvMapCells
		Me.dgMapCells.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgMapCells.Name = "dgMapCells"
		Me.dgMapCells.ReadOnly = True
		Me.dgMapCells.Size = New System.Drawing.Size(944,424)
		Me.dgMapCells.TabIndex = 0
		Me.dgMapCells.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me._
 .DataGridTableStylel})
		'
		'dvMapCells
		'
		Me.dvMapCells.AllowDelete = False
		Me.dvMapCells.AllowEdit = False
		Me.dvMapCells.AllowNew = False
		Me.dvMapCells.ApplyDefaultSort = True
		Me.dvMapCells.Table = Me.dsCiv.MapCell
		'
		'DataGridTableStylel
		'
		Me.DataGridTableStylel.DataGrid = Me.dgMapCells
		Me.DataGridTableStylel.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.DataGridTableStylel.MappingName = ""
		'
		'tabWonders
		'
		Me.tabWonders.Controls.AddRange(New System.Windows.Forms.Control() {Me.dgWonders})
		Me.tabWonders.Location = New System.Drawing.Point(4,22)
		Me.tabWonders.Name = "tabWonders"
		Me.tabWonders.Size = New System.Drawing.Size(952,430)
		Me.tabWonders.TabIndex = 4
		Me.tabWonders.Text = "Wonders"
		Me.tabWonders.Visible = False
		'
		'dgWonders
		'
		Me.dgWonders.DataMember = ""
		Me.dgWonders.DatasSource = Me.dvWonders
		Me.dgWonders.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgWonders.Name = "dgWonders"
		Me.dgWonders.ReadOnly = True
		Me.dgWonders.Size = New System.Drawing.Size(616,424)
		Me.dgWonders.TabIndex = 0
		'
		'tabTriumphs
		'
		Me.tabTriumphs.Controls.AddRange(New System.Windows.Forms.Control()
		Me.tabTriumphs.Location = New System.Drawing.Point(4,22)
		Me.tabTriumphs.Name = "tabTriumphs"
		Me.tabTriumphs.Size = New System.Drawing.Size(952,430)
		Me.tabTriumphs.TabIndex = 1
		Me.tabTriumphs.Text = "Triumphs"
		Me.tabTriumphs.Visible = False
		'
		'dgTriumphs
		'
		Me.dgTriumphs.DataMember = ""
		Me.dgTriumphs.DataSource = Me.dvTriumphs
		Me.dgTriumphs.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgTriumphs.Name = "dgTriumphs"
		Me.dgTriumphs.ReadOnly = True
		Me.dgTriumphs.RowHeadersVisible
		Me.dgTriumphs.Size = New System.Drawing.Size(424,424)
		Me.dgTriumphs.TabIndex = 0
		'
		'tabSummary
		'
		Me.tabSummary.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblMapSeed,_
Me.lblMapSeedTitle, Me.lblVersionTitle, Me.lblVersion, Me.lblTurnsPassedTitle, Me._
lblTurnsPassed, Me.lblYearTitle, Me.lblYear, Me.lblTurnsofPeaceTitle,Me.lblTurnsOfPeace_
, Me.lblDifficultyLevelTitle, Me.lblDifficultylLevel, Me.lblBarbarianActivityTitle, Me._
lblBarbarianActivity, Me.lblNumberOfCitiesTitle, Me.lblNumberOfCities, Me._
lblNumberofUnitsTitle, Me.lblNumberOfUnits, Me.lblCursorHorizTitle, Me.lblCursorHoriz, _
Me.lblMapHeightTitle, Me.lblMapHeight, Me.lblMapWidth, Me.lblMapArea, Me.lblCursorVert, _
Me.LblMapWidthTitle, Me.lblMapAreaTitle, Me.lblCursorVertCoordTitle})
		Me.tabSummary.Location = New System.Drawing.Point(4,22)
		Me.tabSummary.Name = "tabSummary"
		Me.tabSummary.Size = New System.Drawing.Size(952,430)
		Me.tabSummary.TabIndex = 3
		Me.tabSummary.Text = "Summary Info"
		Me.tabSummary.Visible = False
		'
		'lblMapSeed
		'
		Me.lblMapSeed.Location = New System.Drawing.Point(142,192)
		Me.lblMapSeed.Name = "lblMapSeed"
		Me.lblMapSeed.TabIndex = 49
		'
		'lblMapSeedTitle
		'
		Me.lblMapSeedTitle.Location = New System.Drawing.Point(30,392)
		Me.lblMapSeedTitle.Name = "lblMapSeedTitle"
		Me.lblMapSeedTitle.TabIndex = 48
		Me.lblMapSeedTitle.Text = "Map Seed"
		'
		'lblVersionTitle
		'
		Me.lblVersionTitle.Location = New System.Drawing.Point(30,72)
		Me.lblVersionTitle.Name = "lblVersionTitle"
		Me.lblVersionTitle.Size = New System.Drawing.Size(88,23)
		Me.lblVersionTitle.TabIndex = 26
		Me.lblVersionTitle.Text = "Civ Version"
		'
		'lblVersion
		'
		Me.lblVersion.Location = New System.Drawing.Point(30,392)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.Size = New System.Drawing.Size(104,23)
		Me.lblVersion.TabIndex = 22
		'
		'lblTurnsPassedTitle
		'
		Me.lblTurnsPassedTitle.Location = New System.Drawing.Point(254,72)
		Me.lblTurnsPassedTitle.Name = "lblTurnsPassedTitle"
		Me.lblTurnsPassedTitle.Size = New System.Drawing.Size(104,23)
		Me.lblTurnsPassedTitle.TabIndex = 
		Me.lblTurnsPassedTitle.Text = "Turns Passed”
		'
		'IblTurnsPassed
		'
		Me.lblTurnsPassed.Location = New System.Drawing.Point(482,72)
		Me.lblTurnsPassed.Name = "lblTurnsPassed"”
		Me.lblTurnsPassed.Size = New System.Drawing.Size(104,23)
		Me.lblTurnsPassed.TabIndex 23
		'
		'lblYearTitle
		'
		Me.lblYearTitle.Location = New System.Drawing.Point(482,72)
		Me.lblYearTitle.Name ="lblYearTitle" 
		Me.lblYearTitle.Size = New System.Drawing.Size(104,23)
		Me.lblYearTitle.TabIndex = 28
		Me.lblYearTitle.Text = "Year"
		'
		'lblYear
		'
		Me.lblYear.Location = New System.Drawing.Point(597,72)
		Me.lblYear.Name = "lblYear"
		Me.lblYear.Size = New System.Drawing.Size(104,23)
		Me.lblYear.TabIndex = 24
		'
		'lblTurnsofPeaceTitle
		'
		Me.lblTurnsofPeaceTitle.Location = New System.Drawing.Point(712,72) 
		Me.lblTurnsofPeaceTitle.Name = = "lblTurnsofPeaceTitle"
		Me.lblTurnsofPeaceTitle.Size = New System.Drawing.Size(104,23)
		Me.lblTurnsofPeaceTitle.TabIndex  = 32
		Me.lblTurnsofPeaceTitle.Text = "Turns of Peace"
		'
		'lblDifficultylLevelTitle
		'
		Me.lblDifficultylLevelTitle.Location = New System.Drawing.Point(30,232)
		Me.lblDifficultyLevelTitle.Nam e= "lblDifficultyLevelTitle"
		Me.lblDifficultylLevelTitle.Size = New System.Drawing.Size(104,23)
		Me.lblDifficultyLevelTitle.TabIndex = 29
		Me.lblDifficultylLevelTitle.Text= "Difficulty Level"
		'
		'lblDifficultylevel
		'
		Me.lblDifficultyLevel.Location = New System.Drawing.Point(142,232)
		Me.lblDifficultyLevel.Name = "lblDifficultyLevel”
		Me.lblDifficultyLevel.Size = New System.Drawing.Size(104,23)
		Me.lblDifficultyLevel.TabIndex = 25
		'
		'lblBarbarianActivityTitle
		'
		Me.lblBarbarianActivityTitle.Location = New System.Drawing.Point(254,232)
		Me.lblBarbarianActivityTitle.Name = "lblBarbarianActivityTitle"
		Me.lblBarbarianActivityTitle.Size = New System.Drawing.Size(104,23)
		Me.lblBarbarianActivityTitle.TabIndex =30
		Me.lblBarbarianActivityTitle.Text = "Barbarian Activity"
		'
		'lblBarbarianActivity
		'
		Me.lblBarbarianActivity.Location = New System.Drawing.Point(366,232)
		Me.lblBarbarianActivity.Name = "lblBarbarianActivity"
		Me.lblBarbarianActivity.Size = New System.Drawing.Size(104,23)
		Me.lblBarbarianActivity.TabIndex = 31
		'
		'lblNumberOfCitiesTitle
		'
		Me.lblNumberOfCitiesTitle.Location = New System.Drawing.Point(482,232)
		Me.lblNumberOfCitiesTitle.Name =  = "lblNumberOfCitiesTitle"
		Me.lblNumberQfCitiesTitle.Size = New System.Drawing.Size(104,23)
		Me.lblNumberOfCitiesTitle.TabIndex = 33
		Me.lblNumberQfCitiesTitle.Text = "Number of Cities”
		'
		lblNumberOfCities
		'
		Me.lblNumberOfCities.Location = New System.Drawing.Point(597,232)
		Me.lblNumberOfCities.Name = "lblNumberOfCities"
		Me.lblNumberOfCities.Size = New System.Drawing.Size(104,23)
		Me.lblNumberOfCities.TabIndex = 36
		'
		'lblNumberofUnitsTitle
		'
		Me.lblNumberofUnitsTitle.Location = New System.Drawing.Point(712,232)
		Me.lblNumberofUnitsTitle.Name = "lblNumberofUnitsTitle"
		Me.lblNumberofUnitsTitle.Size = New System.Drawing.Size(104,23)
		Me.lblNumberofUnitsTitle.TabIndex = 34
		Me.lblNumberofUnitsTitle.Text = "Numbe of Units"
		'
		'lblNumberOfUnits
		'
		Me.lb1NumberOfUnits.Location = New System.Drawing.Point(827,232)
		Me.lblNumberOfUnits.Name = "lblNumberOfUnits"
		Me.lblNumberOfUnits.Size = New System.Drawing.Size(104,23)
		Me.lblNumberOfUnits.TabIndex = 37
		'
		'lblCursorHorizTitle
		'
		Me.lblCursorHorizTitle.Location = New System.Drawing.Point(712,312)
		Me.lblCursorHorizTitle.Name = "lblCursorHorizTitle"
		Me.lblCursorHorizTitle.Size = New System.Drawing.Size(104,23)
		Me.lblCursorHorizTitle.TabIndex = 44
		Me.lblCursorHorizTitle.Text = "Cursor Horiz Coord"
		'
		'lblCursorHoriz
		'
		Me.lblCursorHoriz.Location = New System.Drawing.Point(827,312)
		Me.lblCursorHoriz.Name = "lblCursorHoriz"
		Me.lblCursorHoriz.Size = New System.Drawing.Size(104,23)
		Me.lblCursorHoriz.TabIndex = 45
		'
		'lblMapHeightTitle
		'
		Me.lblMapHeightTitle.Location = New System.Drawing.Point(30,312)
		Me.lblMapHeightTitle.Name = "lblMapHeightTitle"
		Me.lblMapHeightTitle.Size = New System.Drawing.Size(104,23)
		Me.lblMapHeightTitle.TabIndex = 39
		Me.lblMapHeightTitle.Text = "Map Height"
		'
		'lblMapHeight
		'
		Me.lblMapHeight.Location = New System.Drawing.Point(142,312)
		Me.lblMapHeight.Name =  "lblMapHeight"
		Me.lblMapHeight.Size = New System.Drawing.Size(104,23)
		Me.lblMapHeight.TabIndex = 42
		'
		'lblMapWidth
		'
		Me.lblMapWidth.Location = New System.Drawing.Point(366,312)
		Me.lblMapWidth.Name = "lblMapWidth"
		Me.lblMapWidth.Size = New System.Drawing.Size(104,23)
		Me.lblMapWidth.TabIndex = 41
		'
		'lblMapArea
		'
		Me.lblMapArea.Location = New System.Drawing.Point(597,312) 
		Me.lblMapArea.Name = "lblMapArea"
		Me.lblMapArea.Size = New System.Drawing.Size(104,23)
		Me.lblMapArea.TabIndex = 43
		'
		'lblCursorVert
		'
		Me.lblCursorVert.Location =  New System.Drawing.Point(817,392)
		Me.lblCursorVert.Name = "1blCursorVert"
		Me.lblCursorVert.Size = New System.Drawing.Size(104,23)
		Me.lblCursorVert.TabIndex = 47
		'
		'LblMapWidthTitle
		'
		Me.LblMapWidthTitle.Location = New System.Drawing.Point(254,312)
		Me.LblMapWidthTitle.Name = "LblMapWidthTitle"
		Me.LblMapWidthTitle.Size = New System.Drawing.Size(104,23)
		Me.LblMapWidthTitle.TabIndex = 38
		Me.LblMapWidthTitle.Text = "Map Width"
		'
		'lblMapAreaTitle
		'
		Me.lblMapAreaTitle.Location = New System.Drawing.Size(482,312)
		Me.lblMapAreaTitle.Name = "lblMapAreaTitle"
		Me.lblMapAreaTitle.Size = New System.Drawing.Size(104,23)
		Me.lblMapAreaTitle.TabIndex = 40
		Me.lblMapAreaTitle.Text = "Map Area"
		'
		'lblCursorVertCoordTitle
		'
		Me.lblCursorVertCoordTitle.Location = New System.Drawing.Size(712,392)
		Me.lblCursorVertCoordTitle.Name = "lblCursorVertCoordTitle"”
		Me.lblCursorVertCoordTitle.Size = New System.Drawing.Size(104,23)
		Me.lblCursorVertCoordTitle.TabIndex = 46
		Me.lblCursorVertCoordTitle.Text = "Cursor Vert Coord"
		'
		'tmrSecond
		'
		Me.tmrSecond. Interval = 1000
		'
		'dsCiv
		'
		Me.dsCiv.DataSetName = "Civ"
		Me.dsCiv.EnforceConstraints = False
		Me.dsCiv.Locale = New System.Globalization.CultureInfo("en-GB")
		Me.dsCiv.Namespace = "http://tempuri.org/XMLSchemal.xsd"
		'
		'Forml
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5,13)
		Me.ClientSize = New System.Drawing.Size(976,677)
		Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.tabPages,Me._
 lblOwnerColour, Me.cmbOwnerColour, Me.pnlUnitFilter, Me.grpCounters, Me.pnlCityFilter})
		Me.Name = "Forml"
		Me.Text = "Civ II Snooper"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		CType(Me.dvTriumphs, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvUnits, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvUnitCounts, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvNations, System.ComponentModel.ISupportInitialize).EndInit()
		Me.pnlUnitFilter.ResumeLayout(False)
		Me.grpUnitLocation.ResumeLayout(False)
		CType(Me.dvCities, System.ComponentModel.ISupportInitialize).EndInit()
		Me.grpCounters.ResumeLayout(False)
		CType(Me.dvUnitTypes, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvCivs, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvWonders, System.ComponentModel.ISupportInitialize).EndInit()
		Me.pnlCityFilter.ResumeLayout(False)
		Me.tabPages.ResumeLayout(False)
		Me.tabColSelect.ResumeLayout(False)
		Me.pnlExcel.ResumeLayout(False)
		Me.tabCiv.ResumeLayout(False)
		CType(Me.dgCivs, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabNat.ResumeLayout(False)
		CType(Me.dgNat, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabCities.ResumeLayout(False)
		CType(Me.dgCities, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabUnits.ResumeLayout(False)
		CType(Me.dgUnits, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabUnitCount.ResumeLayout(False)
		CType(Me.dgUnitCounts, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabMapCells.ResumeLayout (False)
		CType(Me.dgMapCells, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.dvMapCells, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabWonders.ResumeLayout (False)
		CType(Me.dgWonders, System.ComponentModel.ISupportInitialize).EndInit()
		Me.tabTriumphs.ResumeLayout(False)
		CType(Me.dgTriumphs,CSystem.ComponentModel.ISupportInitialize).EndInit()
		Me.tabSummary.ResumeLayout(False)
		CType(Me.dsCiv, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Resumelayout(False)
	End Sub
 #End Region
 ===
 #Region "Pre Form Load Initialize
 Private
 Sub Setup/()
 CreateStatusBar()
 SetupColumnStyles
 SetupDBR()
 SetupDGTS
 m
 ()
 ()
 UnitsOwned = False
 '" Setup keys for Dataviews
 dvUnitCounts.Sort
 Component Routines”
 = "UnitType, NationNumber"
 dvUnitTypes.Sort = "UnitType"
 progress.Text = "Ready..."
 End Sub
 Private
 '
 Sub CreateStatusBar
 ()
 Display the first panel with a sunken border style.
 Initialize
 progress.BorderStyle = StatusBarPanelBorderStyle.Sunken
 '
 the text of the panel.
 progress.Text = "Ready..."
 '
 .EndInit
 .EndInit
 ()
 .EndInit
 ()
 .EndInit()
 .EndInit()
 ()
 ()
 Set the AutoSize property to use all remaining space on the StatusBar.
 progress.AutoSize = StatusBarPanelAutoSize.Spring
 '
 Display the second panel with a raised border style.
 time.BorderStyle = StatusBarPanelBorderStyle.Raised
 '
 Create ToolTip text that displays the current time.
 time.ToolTipText
 ()
 ()
 18
 = System.DateTime.Now.ToLongDateString & " " &System.DateTime.Now. «
 ToLongTimeString
 '
 8et the text of the panel to the current date.
 time.Text = System.DateTime.Today.TolLongDateString
 ToLongTimeString
 t
 & " " &System.DateTime.Now.
 Set the AutoSize property to size the panel to the size of the contents.
 time.AutoSize = StatusBarPanelAutoSize.Contents
 '
 Add both panels to the StatusBarPanelCollection
 sb.Panels.Add (progress)
 of the StatusBar.
 v4
 sb.Panels.Add (time)
 '
 Display panels in the StatusBar control.
 sb.ShowPanels = True
 '
 Add the StatusBar
 to the form.
		Me.Controls.Add (sb)
 AddHandler tmrSecond.Tick, AddressOf TmrSecProcessor
 tmrSecond.Start
 End Sub
 Private
 ()
 Sub TmrSecProcessor (Byval ©As Object, ByVal e As EventArgs)
 '
 Create ToolTip text that displays the current time.
 19
 time.ToolTipText = System.DateTime.Now.ToLongDateString & " " & System.DateTime.Now. «
 ToLongTimeString
 '
 Set the text of the panel to the current date.
 time.Text = System.DateTime.Today.ToLongDateString
 ToLongTimeString
 & " " &System.DateTime.Now.
 Set the AutoSize property to size the panel to the size of the contents.
 time.AutoSize = StatusBarPanelAutoSize.Contents
 End Sub
 #Region "Column Styles"
 Private
 Sub SetupColumnStyles()
 Dim r As System.Resources.ResourceManager
 GetType (Forml) )
 SetupNationColumns (r)
 SetupCivColumns
 SetupCityColumns (r)
 SetupUnitColumns (r)
 SetupWonderColumns (r)
 SetupTriumphColumns (r)
 SetupTreatiesColumns (r)
 SetupUnitNationTotalsColumns
 SetupUnitTypeColumns (r)
 End Sub
 #Region "Nations"
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Private
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 NationNumber
 Nation
 Sub SetupNationColumns
		Me.NationNumber
		Me.Nation
 = New System.Resources.ResourceManager
 (xr)
 NationColourNumber
 NationColour
 (x)
 As CIV _II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV IT Extractor.DataGridColouredTextBoxColumn
 NationActive
 As CIV II Extractor.DataGridColouredBoolColumn
 NationHasBeenActive
 NationExtinct
 As CIV II Extractor.DataGridColouredBoolColumn
 As CIV II Extractor.DataGridColouredBoolColumn
 (ByRef resources
 As System.Resources.ResourceManager)
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.NationColourNumber
		Me.NationColour
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV ITExtractor.DataGridColouredTextBoxColumn
		Me.NationActive
		Me.NationHasBeenActive
		Me.NationExtinct
 mNationsCS
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New System.Windows.Forms.DataGridColumnStyle()
 ()
 ()
 ()
 ()
 ()
 {Me.NationNumber,
		Me.Nation, Me.NationColourNumber, Me.NationColour, Me.NationActive,
		Me.NationHasBeenActive, Me.NationExtinct}
 ¥
 'NationNumber
		Me.NationNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.NationNumber.Format
 = ""
		Me.NationNumber.FormatInfo = Nothing
		Me.NationNumber.HeaderText = "Nation#"
		Me.NationNumber.MappingName
 = "NationNumber"
		Me.NationNumber.ReadOnly = True
		Me.NationNumber. Width = 75
 "Nation
 f
 = ""
 ()
 ()
 v4
 (
 "4
		Me.Nation.Format
		Me.Nation.FormatInfo
		Me.Nation.HeaderText
 = Nothing
 = "Nation"
		Me.Nation.MappingName = "Nation"
		Me.Nation.ReadOnly = True
		Me.Nation.Width
 f
 "NationColourNumber
 H
 = 75
 II
 Extractor\Forml.vb
		Me.NationColourNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.NationColourNumber.Format
 = ""
		Me.NationColourNumber.FormatInfo
 = Nothing
		Me.NationColourNumber.HeaderText = "Colour #"
		Me.NationColourNumber.MappingName = "NationColourNumber"
		Me.NationColourNumber.ReadOnly = True
		Me.NationColourNumber. Width = 50
 f
 "NationColour
 ¥
		Me.NationColour.Format
 = ""
		Me.NationColour.FormatInfo
		Me.NationColour.HeaderText
 = Nothing
 = "Nation Colour"
		Me.NationColour.MappingName = "NationColour"
		Me.NationColour.ReadOnly
		Me.NationColour.Width
 = True
 = 75
 13
 "NationActive
 Li
		Me.NationActive.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.NationActive.FalseValue = False
		Me.NationActive.HeaderText = "Active"
		Me.NationActive.MappingName = "NationActive"
		Me.NationActive.NullValue
 System.DBNull)
		Me.NationActive.ReadOnly
		Me.NationActive.TrueValue
		Me.NationActive.Width
 f
 'NationHasBeenlActive
 i
 = CType (resources.GetObject
 = True
 = True
 = 65
		Me.NationHasBeenActive.Alignment
		Me.NationHasBeenActive.FalseValue = False
		Me.NationHasBeenActive.HeaderText
 ("NationActive.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
 = "Was Active"
		Me.NationHasBeenActive.MappingName = "NationHasBeenActive"
		Me.NationHasBeenActive.NullValue
 NullValue"),
 System.DBNull)
 = CType (resources.GetObject
		Me.NationHasBeenActive.ReadOnly = True
		Me.NationHasBeenActive.TrueValue
		Me.NationHasBeenActive.Width
 t
 '‘NationExtinct
 H
		Me.NationExtinct.Alignment
 = 65
 = True
 ("NationHasBeenActive.
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.NationExtinct.FalseValue = False
		Me.NationExtinct.HeaderText = "Extinct"
		Me.NationExtinct.MappingName = "NationExtinect"
		Me.NationExtinct.NullValue
 System.DBNull)
		Me.NationExtinct.ReadOnly
		Me.NationExtinct.TrueValue
		Me.NationExtinct.Width
 End Sub
 #End Region
 #Region "Ci
 Friend
 Friend
 WithEvents
 = CType (resources.GetObject
 = True
 = True
 = 65
 ("NationExtinct.NullValue"),
 CivColour As CIV II Extractor.DataGridColouredTextBoxColumn
 Gender As CIV_II Extractor.DataGridColouredTextBoxColumn
 Friend
 Friend
 Friend
 Friend
 Friend
 Gold As CIV II Extractor.DataGridColouredTextBoxColumn
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 Researching
 ResearchProgress
 SciencePercent
 TaxPercent
 As CIV ITIExtractor.DataGridColouredTextBoxColumn
 As CIV _II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 20
 4
 "4
 WithEvents
C:\Documents and Settings\Jeremy Flowers\My... Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 LuxuryPercent As CIV II Extractor.DataGridColouredTextBoxColumn
 GovernmentType As CIV II Extractor.DataGridColouredTextBoxColumn
 AdvancedFlight As CIV II Extractor.DataGridColouredBoolColumn
 Alphabet As CIV II Extractor.DataGridColouredBoolColumn
 AmphibiousWarfare As CIV II Extractor.DataGridColouredBoolColumn
 Astronomy As CIV II Extractor.DataGridColouredBoolColumn
 AtomicTheoryAsCIVIIExtractor.DataGridColouredBoolColumn Automobile AsCIVITExtractor. DataGridColouredBoolColumn
 BankingAsCIVII_Extractor. DataGridColouredBoolColumn
 BridgeBuilding .As"CIV IT Extractor. DataGridColouredBoolColumn
 BronzeWorking As CIVIIExtractor.DataGridColouredBoolColumn
 CeremonialBurial As CIV II Extractor.DataGridColouredBoolColumn
 Chemistry AsCIVIIExtractor.DataGridColouredBoolColumn
 Chivalry AsCIVIH_Extractor. DataGridColouredBoolColumn
 CodeOfLaws As CIV TT_Extractor.DataGridColouredBoolColumn
 CombinedArms As ETE =
 TT_Extractor.DataGridColouredBoolColumn
 Combustion As CIV £57Extractor. DataGridColouredBoolColumn
 Communism As CIV II Extractor. DataGridColouredBoolColumn
 Computers As CIV II Extractor.DataGridColouredBoolColumn
 Conscription As CIV II Extractor.DataGridColouredBoolColumn
 Construction As CIV II Extractor.DataGridColouredBoolColumn
 Corporation As CIV II Extractor.DataGridColouredBoolColumn
 Currency As CIV IT Extractor.DataGridColouredBoolColumn
 Democracy As CIV II Extractor.DataGridColouredBoolColumn
 Economics As CIVIIExtractor.DataGridColouredBoolColumn
 Electricity As CIVi IT_Extractor.DataGridColouredBoolColumn
 Electronics As CIV IT_Extractor. DataGridColouredBoolColumn
 Engineering As CIV II Extractor.DataGridColouredBoolColumn
 Environmentalism As CIV IIExtractor.DataGridColouredBoolColumn
 Espionage AsCIVIIExtractor.DataGridColouredBoolColumn
 Explosives As CIV IT_Extractor.DataGridColouredBoolColumn FeudalismAsCIVBT_Extractor. DataGridColouredBoolColumn
 Flight As CIV IIExtractor.DataGridColouredBoolColumn
 Fundamentalism As CIV II Extractor.DataGridColouredBoolColumn
 FusionPower As CIV II Extractor.DataGridColouredBoolColumn
 GeneticEngineering As CIV II Extractor.DataGridColouredBoolColumn
 GuerillaWarfare As CIV II Extractor.DataGridColouredBoolColumn
 Gunpowder As CIV IIExtractor.DataGridColouredBoolColumn
 HorsebackRiding ASCTWITExtractor.DataGridColouredBoolColumn
 Industrialization As CIV II Extractor.DataGridColouredBoolColumn
 Invention AsCIVIT_Extractor. DataGridColouredBoolColumn
 IronWorking As CIV II _Extractor.DataGridColouredBoolColumn
 LaborUnionAsCIVIIExtractor. DataGridColouredBoolColumn
 Laser AsCIVIT_Extractor. DataGridColouredBoolColumn
 Leadership As CIV II Extractor.DataGridColouredBoolColumn
 Literacy As CIV II Extractor.DataGridColouredBoolColumn
 MachineTools As“CTV: IIExtractor.DataGridColouredBoolColumn
 MagnetismAsCIVIT Extractor.DataGridColouredBoolColumn
 MapMaking As cing TATE
 _Extractor. DataGridColouredBoolColumn
 Masonry As CIVTTExtractor. DataGridColouredBoolColumn MassProductionAsOTE IIExtractor.DataGridColouredBoolColumn
 Mathematics As CIV IIExtractor.DataGridColouredBoolColumn
		Medecine As CIV II Extractor.DataGridColouredBoolColumn
		Metallurgy As CIV II Extractor.DataGridColouredBoolColumn
 Miniaturization As CIV II Extractor.DataGridColouredBoolColumn
 MobileWarfare As CIVIIExtractor.DataGridColouredBoolColumn
 Monarchy As CIV II Extractor.DataGridColouredBoolColumn
 Monoetheism As CIV II Extractor.DataGridColouredBoolColumn
 Mysticism AsCIVIIExtractor.DataGridColouredBoolColumn
 Navigation As CFV IT_Extractor.DataGridColouredBoolColumn NuclearFission AsOTE ITExtractor.DataGridColouredBoolColumn
 NuclearPower As CIV II Extractor.DataGridColouredBoolColumn
 Philosophy As CIV ITExtractor.DataGridColouredBoolColumn
 Physics As CIV II Extractor.DataGridColouredBoolColumn
 Plastics As CIV II Extractor.DataGridColouredBoolColumn
 Plumbing As CIV IT Extractor.DataGridColouredBoolColumn
 Polytheism As CIV II Extractor.DataGridColouredBoolColumn
 Pottery As CIV II Extractor.DataGridColouredBoolColumn
 Radio As CIV_IIExtractor.DataGridColouredBoolColumn
 Railroad As CIV II Extractor.DataGridColouredBoolColumn
 Recycling As CIV II Extractor.DataGridColouredBoolColumn
 Refining As CIV II EXtractor.DataGridColouredBoolColumn
 Refrigeration As CIV IIExtractor.DataGridColouredBoolColumn
 21
C:\Documents and Settings\Jeremy Flowers\My... Friend WithEvents Republic As CIV II Extractor.DataGridColouredBooclColumn
 Friend WithEvents Robotics As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Rocketry As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Sanitation As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Seafaring As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents SpaceFlight AsCIV_II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Stealth As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents SteamEngine As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Steel As CIV IIExtractor.DataGridColouredBoolColumn
 Friend WithEvents Superconductor As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Tactics As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Theology As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents TheoryOfGravity As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Trade AsCIVIIExtractor.DataGridColouredBoolColumn
 Friend WithEvents University As CIV ITIExtractor.DataGridColouredBoolColumn
 Friend WithEvents WarriorCode As CIV EF_Extractor.DataGridColouredBoolColumn
 Friend WithEvents Wheel As CIV II_Extractor. DataGridColouredBoolColumn
 Friend WithEvents Writing As CIV II Extractor. DataGridColouredBoolColumn
 Private Sub SetupCivColumns (ByRef resources As System.Resources.ResourceManager)
		Me.CivColour = New CIV II Extractor.DataGridColouredTextBoxColumn ()
		Me.Gender = New CIV II Extractor.DataGridColouredTextBoxColumn ()
		Me.Gold = New CIV TF Extractor. DataGridColouredTextBoxColumn ()
		Me. Riss sci = NewWehi'ge ITIExtractor.DataGridColouredTextBoxColumn ()
		Me. Researeli Progress: = New CTVeII Extractor.DataGridColouredTextBoxColumn ()
		Me.SciencePercent = New CIV_II Extractor. DataGridColouredTextBoxColumn ()
		Me.TaxPercent = New CIV II Extractor.DataGridColouredTextBoxColumn ()
		Me.LuxuryPercent = New CIV II Extractor.DataGridColouredTextBoxColumn ()
		Me.GovernmentType = New CIV II Extractor.DataGridColouredTextBoxColumn ()
		Me.AdvancedFlight = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Alphabet =NewCIVII Extractor.DataGridColouredBoolColumn ()
		Me.AmphibiousWarfare = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Astronomy = New CIV IT_Extractor. DataGridColouredBoolColumn ()
		Me:Al TASH = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Automobile = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Banking = New CIV II Extractor. DataGridColouredBoolColumn ()
		Me.BridgeBuilding = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.BronzeWorking = New CIV ITIExtractor.DataGridColouredBoolColumn ()
		Me.CeremonialBurial = New CIV II Extractor.DataGridColouredBoolColumn/()
		Me.Chemistry = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Chivalry = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.CodeOfLaws = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.CombinedArms = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Combustion = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Communism = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Computers = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Conscription = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Construction = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Corporation = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Currency = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Democracy = New CIV ITExtractor.DataGridColouredBoolColumn ()
		Me.Economics = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Electricity = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Electronics = New CIV _ITI Extractor.DataGridColouredBoolColumn ()
		Me.Engineering = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Environmentalism = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Espionage = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Explosives = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Feudalism = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Flight = New CIV ITExtractor.DataGridColouredBoolColumn ()
		Me.Fundamentalism = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.FusionPower = NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.GeneticEngineering = New CIV ITIExtractor.DataGridColouredBoolColumn ()
		Me.GuerillaWarfare = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Gunpowder = New CIV II _Extractor. DataGridColouredBoolColumn ()
		Me.HorsebackRiding = New CIM ITExtractor.DataGridColouredBoolColumn ()
		Me.Industrialization = New CIV _II Extractor. DataGridColouredBoolColumn ()
		Me.Invention = New CIV IT Extractor.DataGridColouredBoolColumn ()
		Me.IronWworking = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.LaborUnion = New CIV II Extractor.DataGridColouredBoolColumn ()
		Me.Laser =NewCIV II Extractor.DataGridColouredBoolColumn ()
		Me.Leadership = New CIV ITExtractor.DataGridColouredBoolColumn ()
		Me.Literacy = New CIV_TI Extractor.DataGridColouredBeoolColumn ()
		Me.MachineTools
 = New CIV II Extractor.DataGridColouredBoolColumn/()
		Me.Magnetism
		Me.MapMaking
		Me.Masonry
		Me.MassProduction
		Me.Mathematics
		Me.Medecine
		Me.Metallurgy
		Me.Miniaturization
		Me.MobileWarfare
		Me.Monarchy
		Me.Monoetheism
		Me.Mysticism
		Me.Navigation
		Me.NuclearFission
		Me.NuclearPower
		Me.Philosophy
		Me.Physics
		Me.Plastics
		Me.Plumbing
		Me.Polytheism
		Me. Pottery
		Me.Radio
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV IT_Extractor.
 DataGridColouredBoolColumn
 = New CIV. ITExtractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV _II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.
 DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn()
 = New CIV II _Extractor.DataGridColouredBoolColumn
 = New CIV _IT_Extractor.
 DataGridColouredBoolColumn
 = New CTW IT _Extractor.DataGridColouredBoolColumn
 = New CIV ITIExtractor.DataGridColouredBoolColumn
 = New CIV EE, _Extractor.
 = New CIV II Extractor.
 DataGridColouredBoolColumn
 DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New GAL II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV ITExtractor.DataGridColouredBoolColumn()
		Me.Railroad
		Me.Recycling
		Me.Refining
		Me.Refrigeration
		Me.Republic
		Me.Robotics
		Me.Rocketry
		Me.Sanitation
		Me.Seafaring
		Me.SpaceFlight
		Me.Stealth
		Me.SteamEngine
		Me.Steel
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn()
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn()
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV ITExtractor.DataGridColouredBoolColumn
 = New CIV II Extractor.
 = New CIV II Extractor.
		Me. Superconductor
 DataGridColouredBoolColumn
 = New EIV ITExtractor.DataGridColouredBoolColumn
 DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.Tactics
		Me.Theology
 = New CIV II _Extractor.
 DataGridColouredBoolColumn
 = New CIV TI_Extractor.DataGridColouredBoolColumn
		Me.TheoryOfGravity
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.Trade
		Me.University
 = New CIV IT_ Extractor.
 .DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.WarriorCode
		Me.Wheel
 = New CIV IT Extractor.
 = New CIV II _Extractor.
 DataGridColouredBoolColumn()
 DataGridColouredBoolColumn
		Me.Writing
 mCivsCS
 = New CIV IT Extractor.DataGridColouredBoolColumn
 = New System.Windows.Forms.DataGridColumnStyle()
		Me.Gender, Me.Gold, Me.Researching, Me.ResearchProgress,
		Me.TaxPercent,
		Me.Alphabet,
		Me.LuxuryPercent,
		Me.AmphibiocusWarfare,
		Me.GovernmentType,
		Me.Astronomy,
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 {)
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 {Me.CivColour,
		Me.ScienceDercent,
		Me.AdvancedFlight,
		Me.AtomicTheory,
		Me.Automobile, Me.Banking, Me.BridgeBuilding, Me.BronzeWorking,
		Me.CeremonialBurial,
		Me.CombinedArms,
		Me.Construction,
		Me.Electricity,
		Me.Chemistry,
		Me.Combustion,
		Me.Corporation,
		Me.Electronics,
		Me.Chivalry,
		Me.CodeOfLaws,
		Me.Communism, Me.Computers,
		Me.Currency,
		Me.Conscription,
		Me.Democracy, Me.Economics,
		Me.Engineering, Me.Environmentalism,
		Me.Espionage, Me.Explosives, Me.Feudalism, Me.Flight, Me. Fundamentalism,
		Me.FusionPower,
		Me.GeneticEngineering,
		Me.HorsebackRiding,
		Me.Industrialization,
		Me.GuerillaWarfare,
		Me.Invention,
		Me.Gunpowder,
		Me.IronWorking,
		Me.LaborUnion, Me.Laser, Me.Leadership, Me.Literacy, Me.MachineTools,
		Me.Magnetism, Me.MapMaking, Me.Masonry, Me.MassProduction,
		Me.Medecine,
		Me.Metallurgy,
		Me.Miniaturization,
		Me.Mathematics,
		Me.MobileWarfare,
 23
 BB
		Me.Monarchyw«
		Me.Monoetheism, Me.Mysticism, Me.Navigation, Me.NuclearFission,
		Me.NuclearPower, Me.Philosophy, Me.Physics, Me.Plastics,
 _
		Me.Plumbing,
 «
		Me.Polytheism, Me.Pottery, Me.Radio, Me.Railroad, Me.Recycling, Me.Refining, «
		Me.Refrigeration,
		Me.Seafaring,
		Me.Republic, Me.Robotics, Me.Rocketry, Me.Sanitation,
		Me.SpaceFlight,
		Me.Superconductor,
		Me.University
		Me.Tactics,
 , Me.WarriorCode,
		Me.Stealth,
		Me.SteamEngine, Me.Steel,
		Me.Theology, Me.TheoryOfGravity, Me.Trade,
		Me.Wheel, Me.Writing)
 "'CivColour
 t
		Me.CivColour.Format
 = ""
		Me.CivColour.FormatInfo
 = Nothing
		Me.CivColour.HeaderText = "Colour"
		Me.CivColour.MappingName = "CivColour"
		Me.CivColour.ReadOnly = True
		Me.CivColour.Width
 f
 "Gender
 f
		Me.Gender.Format
 = 60
 = ""
		Me.Gender.FormatInfo = Nothing
		Me.Gender.HeaderText = "Gender"
		Me.Gender .MappingName = "Gender"
		Me.Gender.ReadOnly = True
		Me.Gender.wWwidth
 f
 "Gold
 = 50
		Me.Gold.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.Gold.Format
 = ""
		Me.Gold.FormatInfo = Nothing
		Me.Gold.HeaderText = "Gold"
		Me.Gold.MappingName = "Gold"
		Me.Gold.ReadOnly = True
		Me.Gold.Width
 1
 'Researching
 H
		Me.Researching.Format
 = 30
 = ""
		Me.Researching.FormatInfo
 = Nothing
		Me.Researching.HeaderText = "Researching"
		Me.Researching.MappingName = "Researching"
		Me.Researching.ReadOnly
 = True
		Me.Researching.Width
 ¥
 'ResearchProgress
 §
 = 125
		Me.ResearchProgress.Alignment
		Me.ResearchProgress.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me. ResearchProgress.FormatInfo
 = Nothing
		Me.ResearchProgress.HeaderText = "Progress"
		Me.ResearchProgress.MappingName = "ResearchProgress"”
		Me.ResearchProgress.ReadOnly
		Me.ResearchProgress.Width
 §
 'SciencePercent
		Me.SciencePercent.Alignment
		Me.SciencePercent.Format
		Me.SciencePercent.FormatInfo
 = ""
 = True
 = 50
 = System.Windows.Forms.HorizontalAlignment.Right
 = Nothing
		Me.SciencePercent.HeaderText = "Science %"
		Me.SciencePercent.MappingName = "SciencePercent"
		Me.SciencePercent.ReadOnly
		Me.SclencePercent.Width
 'TaxPercent
 H
		Me.TaxPercent.Alignment
		Mes TaxBereent.Format.
 = ""
		Me.TaxPercent.FormatInfo
 = True
 = 60
 = System.Windows.Forms.HorizontalAlignment.Right
 = Nothing
		Me. TaxPercent .HeadexrText = "Tax %"
		Me.TaxPercent.MappingName = "TaxPercent"
		Me.TaxPercent.ReadOnly = True
		Me.TaxPercent.Width
 f
 'LuxuryPercent
 ¥
 = 50
		Me.LuxuryPercent.Alignment
		Me.LuxuryPercent.Format
		Me.LuxuryPercent.FormatInfo
		Me.LuxuryPercent.HeaderText
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 I
 Nothing
 = "Luxury
 24
 SS"
Flowers\My...
  Me .LuxuryPercent.MappingName = "LuxuryPercent”
		Me.LuxuryPercent.ReadOnly
		Me.LuxuryPercent.Width
 = True
 = 60
 f
 'GovernmentType
 1
		Me.GovernmentType.Format
 = ""
		Me.GovernmentType.FormatInfo = Nothing
		Me.GovernmentType.HeaderText
 = "GovernmentType"
		Me. GovernmmentType.MappingName = "GovernmentType"
		Me.GovernmentType.ReadOnly = True
		Me.GovernmentType.Width
 = 105
 13
 'AdvancedFlight
		Me.AdvancedFlight.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.AdvancedFlight.FalseValue = False
		Me.AdvancedFlight.HeaderText
 "AF"
		Me.AdvancedFlight.MappingName = "AdvancedFlight"
		Me.AdvancedFlight.NullValue
 System.DBNull)
		Me.AdvancedFlight.ReadOnly
		Me.AdvancedFlight.TrueValue
		Me.AdvancedFlight.Width
 "Alphabet
 = CType (resources.GetObject
 = True
 = True
 = 25
		Me.Alphabet.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Alphabet.FalseValue = False
		Me.Alphabet.HeaderText
 = "Alp"
		Me.Alphabet.MappingName = "Alphabet"
		Me.Alphabet.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Alphabet.ReadOnly = True
		Me.Alphabet.TrueValue
		Me.Alphabet.Width
 §
 *AmphibicusWarfare
 ¥
 = 25
 = True
		Me.AmphibiousWarfare.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.AmphibiocusWarfare.FalgeValue = False
		Me.AmphibiousWarfare.HeaderText
		Me.AmphibiocusWarfare.MappingName
		Me.AmphibiousWarfare.NullValue
 = "A W"
 = "AmphibiousWarfare"
 = CType (resources.GetObject
 Nullvalue"),
 System.DBNull)
		Me.AmphibiousWarfare.ReadOnly = True
		Me.AmphibiousWarfare.TrueValue
 = True
		Me.AmphibiousWarfare.Width
 ¥
 "Astronomy
 EH
 = 25
		Me.Astronomy.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Astronomy.FalseValue = False
		Me.Astronomy.HeaderText = "Ast"
		Me.Astronomy.MappingName
		Me.Astronomy.NullvValue
 DBNull)
 = "Astronomy"
 = CType (resources.GetObject
		Me.Astronomy.ReadOnly = True
		Me.Astrononmy.TrueValue = True
		Me.Astronomy.Width
 §
 *AtomicTheory
 f
 = 25
		Me.AtomicTheory.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.AtomicTheory.FalseValue
		Me.AtomicTheory.HeaderText
		Me.AtomicTheory.MappingName
		Me.AtomicTheory.NullValue
 System.DBNull)
 = False
 = "A T"
 = "AtomicTheory"
 = CType (resources.GetObject
		Me.AtomicTheory.ReadOnly = True
		Me.AtomicTheory.TrueValue = True
		Me.AtomicTheory.Width
 25
 ("AdvancedFlight.NullvValue"),
 ("Alphabet.Nullvalue"),
 System.
 ("AmphibiousWarfare.
 ("Astronomy.Nullvalue"),
 ("AtomicTheory.NullvValue"),
 System.
 ¢
 v4
 vs
 «
 v
 = 25
 ‘Automobile
 4
		Me.Automobile.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Automobile.FalseValue = False
		Me.Automobile.HeaderText
 = "Aut"
		Me.Automobile.MappingName = "Automobile"
		Me.Automobile.Nullvalue
 = CType (resources.GetObject
 26
 ("Automobile .NullValue™), System. ¢
 DBNull)
		Me.Automobile.ReadOnly = True
		Me.Automobile.,TrueValue = True
		Me.Automobile.Width
 ¥
 "Banking
 = 25
		Me.Banking.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Banking.FalseValue = False
		Me.Banking.HeaderText
 H
 "Ban"
		Me.Banking.MappingName = "Banking"
		Me.Banking.NullValue
 = CType (resources.GetObject
		Me.Banking.ReadOnly = True
		Me.Banking.TrueValue = True
		Me.Banking.Width
 f
 'BridgeBuilding
 = 25
		Me.BridgeBuilding.Alignment
 ("Banking.NullvValue"),
 System.DBNullw
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.BridgeBuilding.FalseValue
		Me.BridgeBuilding.HeaderText
 False
 = "B B"
		Me.BridgeBuilding.MappingName = "BridgeBuilding"
		Me.BridgeBuilding.NullValue
 System.DBNull)
		Me.BridgeBuilding.ReadOnly
		Me.BridgeBuilding.TrueValue
		Me.BridgeBuilding.Width
 H
 '‘BronzeWorking
 ¥
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("BridgeBuilding.Nullvalue"),
		Me.BronzeWorking.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.BronzeWorking.FalseValue
		Me.BronzeWorking.HeaderText
		Me.BronzeWorking.MappingName
		Me.BronzeWorking.NullValue
 System.DBNull)
 = False
 "B W"
 = "BronzeWorking"
 = CType (resources.GetObject
		Me.BronzeWorking.ReadOnly = True
		Me.BronzeWorking.TrueValue = True
		Me.BronzeWorking.Width
 tf
 '‘CeremonialBurial
 ¥
 = 25
 ("BronzeWorking.NullValue"),
		Me.CeremonialBurial.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.CeremonialBurial.FalseValue = False
		Me.CeremonialBurial .HeaderText = "C B"
		Me.CeremonialBurial
 .MappingName = "CeremonialBurial"
		Me.CeremonialBurial.NullValue
 Nullvalue™), System.DBNull)
		Me.CeremonialBurial.ReadOnly
		Me.CeremonialBurial.TrueValue
		Me.CeremonialBurial.Width
 f
 "Chemistry
 t
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("CeremonialBurial.
		Me.Chemistry.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Chemistry.FalseValue = False
		Me.Chemistry.HeaderText
 = "Che"
		Me.Chemistry.MappingName = "Chemistry"
		Me.Chemistry.NullvValue
 DBNull)
 = CType (resources.GetObject
		Me.Chemistry.ReadOnly = True
		Me.Chemistry.TrueValue
		Me.Chemistry.Width
 t
 = 25
 = True
 ("Chemistry.NullValue”),
 System.
 «¢
 «
 4
 ‘Chivalry
 §
		Me.Chivalry.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Chivalry.FalseValue = False
		Me.Chivalry.HeaderText
 = "Chi"
		Me.Chivalry.MappingName = "Chivalry"
		Me.Chivalry.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Chivalry.ReadOnly = True
		Me.Chivalry.TrueValue
 = True
		Me.Chivalry.Width
 §
 'CodeOfLaws
 ¥
 = 25
 ("Chivalry.NullvValue"),
		Me.CodeOfLaws.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.CodeOfLaws.FalseValue = False
		Me.CodeOflLaws.HeaderText
 = "COL"
		Me.CodeOfLaws .MappingName = "CodeOfLaws"
		Me.CodeOfLaws.NullvValue
 DBNull)
 = CType (resources.GetObject
		Me.CodeOflLaws.ReadOnly = True
		Me.CodeOfLaws.TrueValue = True
		Me.CodeOfLaws.Width
 |
 'CombinedArms
 H
 = 25
 ("CodeOfLaws.NullValue™),
		Me.CombinedArms.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.CombinedArms.FalseValue = False
		Me.CombinedArms.HeaderText
 = "C A"
		Me.CombinedArms.MappingName = "CombinedArms"
		Me.CombinedArms.NullValue
 = CType (resources.GetObject
 System.
 ("CombinedArms.NullValue"),
 System.DBNull)
		Me.CombinedArms.ReadOnly
 = True
		Me.CombinedArms.TrueValue = True
		Me.CombinedArms.Width
 ‘Combustion
 i
 = 25
		Me.Combustion.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Combustion.FalseValue = False
		Me.Combustion.HeaderText
		Me.Combustion.MappingName
		Me.Combustion.NullValue
 DBNull)
 = "Cmb"
 = "Combustion"
 = CType (resources.GetObject
		Me.Combustion.ReadOnly = True
		Me.Combustion.TrueValue = True
		Me.Combustion.Width
 t
 "Communism
 7
 = 25
 ("Combustion.NullValue"),
		Me.Communism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Communism.FalseValue = False
		Me.Communism.HeaderText
		Me,.Communism.MappingName
		Me.Communism.NullValue
 DBNull)
 = "Com"
 = "Communism"
 = CType (resources.GetObject
		Me.Communism.ReadOnly = True
		Me.Communism.TrueValue = True
		Me.Communism.Width
 §
 "Computers
 = 25
 ("Communism.NullValue"),
		Me.Computers.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Computers,FalseValue = False
		Me.Computers.HeaderText
		Me.Computers.MappingName
		Me.Computers.NullValue
 'DBNull)
 = "Cmp"
 = "Computers"
 = CType (resources.GetObject
		Me.Computers.ReadOnly = True
		Me.Computers.TrueValue = True
		Me.Computers.Width
 tf
 '*Conscription
 ¥
 = 25
 ("Computers.NullValue"),
 |
 System.
 System.
 27
 »
 System. wv
 v4
 System. «
 v4
 v4
		Me.Conscription.Alignment
 II
 Extractor\Forml.vb
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Conscription.FalseValue = False
		Me.Conscription.HeaderText
 = "Csc”
		Me.Conscription.MappingName = "Conscription"
		Me.Conscription.NullvValue
 System.DBNull)
		Me.Conscription.ReadOnly
		Me.Conscription.TrueValue
		Me.Conscription.Width
 t
 "Construction
 t
		Me.Construction.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Conscription.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Construction.FalseValue = False
		Me.Construction.HeaderText
 = "Cst"
		Me.Construction.MappingName = "Construction"
		Me.Construction.NullValue
 System.DBNull)
		Me.Construction.ReadOnly
		Me.Construction.TrueValue
		Me.Construction.Width
 EH
 "Corporation
 t
		Me.Corporation.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Construction.Nullvalue™),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Corporation.FalseValue = False
		Me.Corporation.HeaderText
 = "Cor"
		Me.Corporation.MappingName = "Corporation"
		Me.Corporation.Nullvalue
 System.DBNull)
		Me.Corporation.ReadOnly
		Me.Corporation.TrueValue
		Me.Corporation.Width
 H
 ‘Currency
 f
		Me.Currency.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Corporation.NullValue"),
 .
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Currency.FalseValue = False
		Me.Currency.HeaderText
		Me.Currency.MappingName
		Me.Currency.NullValue
 DBNull)
 = "Cur"
 = "Currency"
 = CType (resources.GetObject
		Me.Currency.ReadOnly = True
		Me.Currency.TrueValue = True
		Me.Currency.Width
 H
 "Democracy
 ¥
 = 25
 ("Currency.Nullvalue™),
		Me.Democracy.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Democracy.FalseValue = False
		Me.Democracy.HeaderTexXxt
 = "Dem"
		Me.Democracy.MappingName = "Democracy"
		Me.Democracy.NullValue
 DBN
 ull)
 = CType (resources.GetObject
		Me.Democracy.ReadOnly = True
		Me.Democracy.TrueValue = True
		Me.Democracy.Width
 f
 "Economics
 = 25
 ("Democracy.NullValue”),
		Me.Economics.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Economics.FalseValue = False
		Me.Economics.HeaderText
		Me.Economics.MappingName
		Me.Economics.NullValue
 DBNull)
 = "Eco"
 = "Economics"
 = CType (resources.GetObject
		Me.Economics.ReadOnly = True
		Me.Economics.TrueValue = True
		Me.Economics.Width
 :
 "Electricity
 ¥
 = 25
		Me.Electricity.Alignment
 ("Economics.NullValue"),
 System.
 System.
 System.
 28
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Electricity.FalseValue
		Me.Electricity.HeaderText
		Me.Electricity.MappingName
		Me.Electricity.NullValue
 System.DBNull)
		Me.Electricity.ReadOnly
		Me.Electricity.TrueValue
		Me.Electricity.Width
 ¥
 "Electronics
		Me.Electronics.Alignment
 = 25
 = False
 = "E1t"
 = "Electricity"
 = CType (resources.GetObject
 ("Electricity.NullValue™),
 = True
 = True
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Electronics.FalseValue = False
		Me.hlectronics.HeaderText
		Me.Electronics.MappingName
		Me.Electronics.NullValue
 System.DBNull)
		Me.Electronics.ReadOnly
 = "Eln"
 = "Electronics™
 = CType (resources.GetObject
 = True
		Me.Electronics.TrueValue
		Me.Electronics.Width
 ¥
 "Engineering
 = True
 = 25
 ("Electronics.NullValue"),
		Me.Engineering.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Engineering.FalseValue = False
		Me.Engineering.HeaderText
 = "Eng"
		Me. Engineering .MappingName = "Engineering"
		Me.Engineering.NullValue
 System.DBNull)
		Me.Engineering.ReadOnly
		Me.Engineering.TrueValue
		Me.Engineering.width
 ‘Environmentalism
 L
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Engineering.NullvValue"),
		Me.Environmentalism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Environmentalism.FalseValue = False
		Me.Environmentalism.HeaderText
 = "Env"
		Me.Environmentalism.MappingName = "Environmentalism"
		Me.Environmentalism.NullValue
 NullValue"),
 System.DBNull)
		Me.Environmentalism.ReadOnly
		Me.Environmentalism.TrueValue
		Me.Environmentalism.Width
 "Espionage
 14
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Environmentalism.
		Me.Espionage.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Espionage.FalseValue = False
		Me.Espionage.HeaderText
 = "Esp"
		Me.Espionage.MappingName = "Espionage"
		Me.Espionage.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Espionage.ReadOnly = True
		Me.Espionage.TrueValue
		Me.Espionage.Width
 H
 Explosives
 H
 = 25
 = True
 ("Espionage.NullValue"),
		Me.Explosives.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Explosives.FalseValue = False
		Me.Explosives.HeaderText
 = "Exp"
		Me.Explosives.MappingName = "Explosives"
		Me.Explosives.NullValue
 DBNull)
		Me.Explosives.ReadOnly
		Me.Explosives.TrueValue
		Me.Explosives.Width
 "Feudalism
 H
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Explosives.Nullvalue"),
		Me.Feudalism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
 System.
 29
 4
 v4
 v4
 v4
 v4
 System. «
		Me.Feudalism.FalseValue = False
		Me.Feudalism.HeaderText
 = "Feu"
		Me.Feudalism.MappingName = "Feudalism"
		Me.Feudalism.NullValue
 = CType (resources.GetObject
 DBNull)
		Me.Feudalism.ReadOnly = True
		Me.Feudalism.TrueValue
		Me.Feudalism.Width
 "Flight
		Me.Flight.Alignment
 = 25
 = True
 ("Feudalism.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Flight.FalseValue = False
		Me.Flight.HeaderText
 = "F1i"
		Me.Flight .MappingName = "Flight"
		Me.Flight.NullValue
 = CType (resources.GetObject("Flight.NullValue"),
		Me.Flight.ReadOnly = True
		Me.Flight.TrueValue
		Me.Flight.Width
 '*Fundamentalism
 = 25
 = True
		Me.Fundamentalism.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Fundamentalism.FalseValue = False
		Me.Fundamentalism.HeaderText
 = "Fun"
		Me.Fundamentalism.MappingName = "Fundamentalism"
		Me.Fundamentalism.NullValue
 = CType (resources.GetObject
 System.
 System.DBNull)
 ("Fundamentalism.NullValue"),
 System.DBNull)
		Me. Fundamentalism. ReadOnly = True
		Me. Fundamentalism.
		Me. Fundamentalism.Width
 iH
 "FugionPower
 TrueValue
 = 25
 = True
		Me.FusionPower.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.FusionPower.FalseValue = False
		Me.FusionPower.HeaderText
		Me.FusionPower.MappingName
		Me.FusionPower.NullValue
 System.DBNull)
 = "Fus"
 = "FusionPower"
 = CType (resources.GetObject
		Me.FusionPower.ReadOnly = True
		Me.FusionPower.TrueValue = True
		Me. FusionPower. Width = 25
 4
 'GeneticEngineering
 tf
		Me.GeneticEngineering.Alignment
 ("FusionPower.NullvValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.GeneticEngineering.FalseValue = False
		Me.GeneticEngineering.HeaderText
		Me.GeneticEngineering.MappingName
		Me.GeneticEngineering.NullValue
 NullvValue™), System.DBNull)
		Me.GeneticEngineering.ReadOnly
 = "G E"
 = "GeneticEngineering"
 = CType (resources.GetObject
 = True
		Me.GeneticEngineering.TrueValue
		Me.GeneticEngineering.Width
 §
 "GuerillaWarfare
 = True
 = 25
 ("GeneticEngineering.
		Me.GuerillaWarfare.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.GuerillaWarfare.FalseValue = False
		Me.GuerillaWarfare.HeaderText
 = "G W"
		Me.GuerillaWarfare.MappingName = "GuerillaWarfare"
		Me.GuerillaWwarfare.NullValue
 )
 System.DBNull)
		Me.GuerillaWarfare.ReadOnly
		Me.GuerillaWarfare.TruevValue
		Me.GuerillaWarfare.Width
 §
 ‘Gunpowder
 = CType (resources.GetObject("GuerillaWarfare.NullValue"
 = True
 = True
 = 25
		Me.Gunpowder.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Gunpowder.FalseValue = False
		Me. Gunpowder .HeaderText
 = "Gun"
		Me. Gunpowder .MappingName = "Gunpowder"
 30
 «
 ¢
 "4
 "4
 v
		Me.Gunpowder.NullValue
 DBNull)
 = CType (resources.GetObject
 ("Gunpowder.NullvValue"),
		Me. Gunpowder .ReadOnly = True
		Me.Gunpowder.TrueValue = True
		Me. Gunpowder .Width = 25
 H
 'HorsebackRiding
 §
		Me.HorsebackRiding.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.HorsebackRiding.FalseValue = False
		Me.HorsebackRiding.HeaderText
 = "H R"
		Me.HorsebackRiding.MappingName = "HorsebackRiding"
		Me.HorsebackRiding.NullValue
 = CType (resources.GetObject
 System.
 ("HorsebackRiding.NullValue"
 ),
 System.DBNull)
		Me.HorsebackRiding.ReadOnly = True
		Me.HorsebackRiding.TrueValue
		Me.HorsebackRiding.Width
 H
 "Industrialization
 f
 = 25
		Me.Industrialization.Alignment
		Me.Industrialization.FalseValue
		Me.Industrialization.HeaderText
 = True
 = System.Windows.Forms.HorizontalAlignment.Center
 = False
 = "Ind"
		Me.Industrialization.MappingName = "Industrialization”
		Me.Industrialization.NullValue
 NullValue"),
 System.DBNull)
		Me.Industrialization.ReadOnly
		Me.Industrialization.TrueValue
		Me.Industrialization.Width
 f
 ‘Invention
		Me.Invention.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Industrialization.
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Invention.FalseValue = False
		Me.Invention.HeaderText
 = "Inv"
		Me.Invention.MappingName = "Invention"
		Me.Invention.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Invention.ReadOnly = True
		Me.Invention.TrueValue
		Me.Invention.Width
 'IronWorking
 = 25
 = True
 ("Invention.NullValue"),
 .
		Me.IronWorking.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.IronWorking.FalseValue
 = False
		Me.IronWorking.HeaderText
		Me.IronWorking.MappingName
		Me.IronWorking.NullValue
 System.DBNull)
 = "I W"
 = "IronWorking"
 = CType (resources.GetObject
		Me.IronWorking.ReadOnly = True
		Me.IronWorking.TrueValue
		Me.IronWorking.wWidth
 §
 'LaborUnion
 = 25
 = True
 ("IronWorking.NullvValue"),
		Me.LaborUnion.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.LaborUnion.FalseValue
 = False
		Me.LaborUnion.HeaderText
		Me.LaborUnion.MappingName
		Me.LaborUnion.NullValue
 DBNull)
 = "L U"
 = "LaborUnion"
 = CType (resources.GetObject
		Me.LaborUnion.ReadOnly = True
		Me.LaborUnion.TrueValue = True
		Me.LaborUnion.width
 ‘Laser
 ¥
 = 25
 ("LaborUnion.NullValue"),
		Me.Laser.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Laser.FalseValue = False
		Me.Laser.HeaderText
 = "Las"
		Me.Laser.MappingName = "Laser"
		Me.Laser.NullValue
 = CType (resources.GetObject
 ("Laser.NullValue"),
 System.
 31
 «
 «
 V4
 4
 v4
 System. w«
 System.DBNull)
		Me.Laser.ReadOnly = True
		Me.Laser.TrueValue = True
		Me.Laser.Width
 §
 ‘Leadership
 ¥
 = 25
		Me.Leadership.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Leadership.FalseValue = False
		Me.Leadership.HeaderText
 = "Lea"
		Me.Leadership.MappingName = "Leadership"
		Me.Leadership.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Leadership.ReadOnly = True
		Me.Leadership.TrueValue
 = True
		Me.Leadership.Width
 §
 "Literacy
 H
		Me.Literacy.Alignment
 = 25
 ("Leadership.NullValue”),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Literacy.FalseValue = False
		Me.Literacy.HeaderText = "Lit"
		Me.Literacy.MappingName = "Literacy"
		Me.Literacy.NullValue
 DBNull)
		Me.Literacy.ReadOnly
		Me.Literacy.TrueValue
		Me.Literacy.Width
 §
 'MachineTools
 f
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Literacy.NullvValue"),
		Me.MachineTools.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.MachineTools.FalseValue = False
		Me.MachineTools.HeaderText
 = "M T"
		Me.MachineTools.MappingName = "MachineTools"
		Me.MachineTools.NullValue
 = CType (resources.GetObject
 System.
 ("MachineTools.NullvValue"),
 System.DBNull)
		Me.MachineTools.ReadOnly = True
		Me.MachineTools.TrueValue = True
		Me.MachineTools.Width
 {
 "Magnetism
 H
 = 25
		Me.Magnetism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Magnetism.FalseValue
		Me.Magnetism.HeaderText
 = False
 = "Mag"
		Me.Magnetism.MappingName = "Magnetism"
		Me.Magnetism.NullValue
 DBNull).
 = CType (resources.GetObject
		Me.Magnetism.ReadOnly = True
		Me.Magnetism.TrueValue = True
		Me.Magnetism.Width
 ¥
 ‘*MapMaking
 i
 = 25
 ("Magnetism.NullValue”),
		Me.MapMaking.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.MapMaking.FalseValue = False
		Me.MapMaking.HeaderText
 = "M M"
		Me.MapMaking.MappingName = "MapMaking"
		Me.MapMaking.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.MapMaking.ReadOnly = True
		Me.MapMaking.TrueValue = True
		Me.MapMaking.Width
 "Masonry
 = 25
 ("MapMaking.NullValue"),
		Me.Masonry.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Masonry.FalseValue = False
		Me.Masonry.HeaderText
 = "Mas"
		Me.Masonry.MappingName = "Masonry"
		Me.Masonry.NullValue
 = CType (resources.GetObject
 ("Masonry.NullValue"),
 System.
 System.
 System.DBNull
 32
 System. «
 v4
 v4
 V4
 4
		Me.Masonry.ReadOnly = True
 «
		Me.Masonry.TrueValue = True
		Me.Masonry.Width
 'MassProduction
 LH
 = 25
		Me.MassProduction.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.MassProduction.FalseValue = False
		Me.MassProduction.HeaderText
 = "M P"
		Me.MassProduction.MappingName = "MassProduction"
		Me.MassProduction.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.MassProduction.ReadOnly = True
		Me.MassProduction.TrueValue
		Me.MassProduction.Width
 ‘Mathematics
 H
		Me.Mathematics.Alignment
 = 25
 = True
 ("MassProduction.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Mathematics.FalseValue = False
		Me.Mathematics.HeaderText
		Me.Mathematics.MappingName
 |
 = "Mat"
 = "Mathematics™
		Me.Mathematics.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.Mathematics.ReadOnly = True
		Me.Mathematics.TrueValue
		Me.Mathematics.Width
 §
 "Medecine
 ¥
 = 25
 = True
 ("Mathematics.NullValue"),
		Me.Medecine.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Medecine.FalseValue = False
		Me.Medecine.HeaderText
		Me.Medecine.MappingName
 = "Med"
 = "Medecine"
		Me.Medecine.NullValue = CType(resources.GetObject
 DBNull)
		Me.Medecine.ReadOnly = True
		Me.Medecine.TrueValue = True
		Me.Medecine.Width
 H
 ‘Metallurgy
 ¥
 = 25
		Me.Metallurgy.Alignment
 ("Medecine.NullvValue"), System.
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Metallurgy.FalseValue = False
		Me.Metallurgy.HeaderText
 = "Met"
		Me.Metallurgy.MappingName = "Metallurgy"
		Me.Metallurgy.NullValue
 DBNull)
		Me.Metallurgy.ReadOnly
		Me.Metallurgy.TrueValue
		Me.Metallurgy.Width
 f
 'Miniaturization
 {
 = CType (resources.GetObject
 = True
 = True
 = 25
		Me.Miniaturization.Alignment
		Me.Miniaturization.FalseValue
 ("Metallurgy.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
 = False
		Me.Miniaturization.HeaderText
		Me.Miniaturization.MappingName
		Me.Miniaturization.NullValue
 )
 System.DBNull)
		Me.Miniaturization.ReadOnly
		Me.Miniaturization.TrueValue
		Me.Miniaturization.wWwidth
 §
 'MoblileWarfare
		Me.MobileWarfare.Alignment
 = 25
 = "Min"
 = "Miniaturization"
 = (Type (resources.GetObject
 = True
 = True
		Me.MobileWarfare.FalseValue = False
 ("™Minlaturization.NullValue"
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.MobileWarfare.HeaderText
		Me.MobileWarfare.MappingName
		Me.MobileWarfare.NullValue
 System.DBNull)
 = "M W"
 = "MobileWarfare™
 = CType (resources.GetObject
		Me.MobileWarfare.ReadOnly = True
		Me.MobileWarfare.TrueValue
 ("MobileWarfare.NullValue"),
 33
 ¢
 v4
 v4
 System. wv
 v
 vs
 = True
		Me.MobileWarfare.Width
 "Monarchy
 §
 = 25
		Me.Monarchy.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Monarchy.FalseValue = False
		Me.Monarchy.HeaderText
 = "Mrc"
		Me.Monarchy.MappingName = "Monarchy"
		Me.Monarchy.NullvValue
 DBNull)
 = CType (resources.GetObject
		Me.Monarchy.ReadOnly = True
		Me.Monarchy.TrueValue = True
		Me.Monarchy.Width
 {
 'Monoetheism
 1
 = 25
 ("Monarchy.NullValue"),
		Me.Monoetheism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Monoetheism.FalseValue
		Me.Monoetheism.HeaderText
		Me.Monoetheism.MappingName
		Me.Monoetheism.NullValue
 = False
 = "Moe"
 = "Monotheism"
 = CType (resources.GetObject
 System.
 ("Monoetheism.NullValue"),
 System.DBNull)
		Me.Monoetheism.ReadOnly = True
		Me.Monoetheism.TrueValue = True
		Me.Monoetheism.Width
 = 25
		Me.Mysticism.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Mysticism.FalseValue = False
		Me.Mysticism.HeaderText
 = "Mys"
		Me.Mysticism.MappingName = "Mysticism"
		Me.Mysticism.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Mysticism.ReadOnly = True
		Me.Mysticism.TrueValue = True
		Me.Mysticism.Width
 "Navigation
 f
 = 25
		Me.Navigation.Alignment
 ("Mysticism.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Navigation.FalseValue = False
		Me.Navigation.HeaderText
 = "Nav"
		Me.Navigation.MappingName = "Navigation"
		Me.Navigation.NullValue
 = CType (resources.GetObject
 ("Navigation.Nullvalue"),
 System.
 34
 Vs
 v4
 «¢
 System. ¢
 DBNull)
		Me.Navigation.ReadOnly
		Me.Navigation.TrueValue
		Me.Navigation.Width
 = True
 = True
 = 25
 ¥
 'NuclearFission
 .
		Me.NuclearFission.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.NuclearFission.FalseValue = False
		Me.NuclearFission.HeaderText
 = "N FE"
		Me.NuclearFission.MappingName = "NuclearFission"
		Me.NuclearFission.NullValue
 System.DBNull)
		Me.NuclearFission.ReadOnly
		Me.NuclearFission.TrueValue
		Me.NuclearFission.Width
 {
 '‘NuclearPower
 ¥
		Me.NuclearPower.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("NuclearFission.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.NuclearPower.FalsevValue = False
		Me.NuclearPower.HeaderText
 = "N P"
		Me.NuclearPower.MappingName = "NuclearPower"
		Me.NuclearPower.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.NuclearPower.ReadOnly = True
		Me.NuclearPower.TrueValue
 = True
		Me.NuclearPower.Width
 = 25
 ("NuclearPower.Nullvalue"),
 «
 ve
 4
 "Philosophy
		Me.Philosophy.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Philosophy.FalseValue = False
		Me.Philosophy.HeaderText
 = "Phi"
		Me.Philosophy.MappingName = "Philosophy"
		Me.Philosophy.NullValue
 DBNull)
		Me.Philosophy.ReadOnly
		Me.Philosophy.TrueValue
		Me.Philosophy.Width
 ‘Physics
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Philosophy.NullvValue”),
		Me.Physics.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Physics.FalseValue = False
		Me.Physics.HeaderText
 = "Phy"
		Me.Physics.MappingName = "Physics"
		Me.Physics.NullValue
 = CType (resources.GetObject
 ("Physics.NullValue"),
 35
 System. «
 System.DBNullw
		Me.Physics.ReadOnly = True
		Me.Physics.TrueValue
		Me.Physics.Width
 ;
 "Plastics
 §
 = 25
 = True
		Me.Plastics.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Plastics.FalseValue = False
		Me.Plastics.HeaderText = "Pla"
		Me.Plastics.MappingName = "Plastics"
		Me.Plastics.NullValue
 DBNull)
		Me.Plastics.ReadOnly
		Me.Plastics.TrueValue
		Me.Plastics.Width
 ¥
 "Plumbing
 1
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Plastics.NullValue"),
		Me.Plumbing.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Plumbing.FalseValue = False
		Me.Plumbing.HeaderText = "Plu"
		Me.Plumbing.MappingName
 = "Plumbing"
		Me.Plumbing.NullValue
 DBN
 ull)
 = CType (resources.GetObject
		Me.Plumbing.ReadOnly = True
		Me.Plumbing.TrueValue = True
		Me.Plumbing.Width
 t
 'Polytheism
 t
 = 25
		Me.Polytheism.Alignment
 ("Plumbing.NullValue"”),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Polytheism.FalseValue = False
		Me.Polytheism.HeaderText
 = "Pol"
		Me.Polytheism.MappingName = "Polytheism"
		Me.Polytheism.NullValue
 = CType (resources.GetObject
 ("Polytheism.NullValue”),
 System.
 System.
 v4
 v
 System. «
 DBNull)
		Me.Polytheism.ReadOnly = True
		Me.Polytheism.TrueValue
		Me.Polytheism.Width
 ‘Pottery
 f
 = 25
 = True
		Me.Pottery.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Pottery.FalseValue = False
		Me.Pottery.HeaderText
 = "Pot"
		Me.Pottery.MappingName = "Pottery"
		Me.Pottery.NullValue
		Me.Pottery.ReadOnly
		Me. Pottery.TrueValue
		Me.Pottery.Width
 = CType (resources.GetObject
 = True
 = True
 ("Pottery.NullValue"),
 System.DBNull
 ww
 = 235
 Radio
 ¥f
		Me.Radio.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Radio.FalseValue
		Me.Radio.HeaderText
 I
 i
 False
 "Rad"
		Me.Radio.MappingName = "Radio"
		Me.Radio.NullValue
 = CType (resources.GetObject
		Me.Radio.ReadOnly = True
		Me.Radio.TrueValue = True
		Me.Radio.Width
 F
 "Railroad
 = 23
		Me.Railroad.Alignment
 ("Radio.NullvValue"),
 System.DBNull)
 = System.Windows.Forms. HorizontalAlignment.Center
		Me.Railroad.FalseValue = False
		Me.Railroad.HeaderText
 = "Rai"
		Me.Railroad.MappingName = "Railroad"
		Me.Railroad.NullValue
 DBNull)
		Me.Railroad.ReadOnly
		Me.Railroad.TrueValue
		Me.Railroad.Width
 t
 "Recycling
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("Railrocad.NullValue”),
		Me.Recycling.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Recycling.FalseValue = False
		Me.Recycling.HeaderText
 = "Rec"
		Me.Recycling.MappingName = "Recycling"
		Me.Recycling.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Recycling.ReadOnly = True
		Me.Recvyecling.TrueValue = True
		Me.Recycling.Width
 f
 "Refining
 H
		Me.Refining.Alignment
 = 25
 ("Recycling.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Refining.FalseValue = False
		Me.Refining.HeaderText
 = "Rfn"
		Me.Refining.MappingName = "Refining"
		Me.Refining.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Refining.ReadOnly = True
		Me.Refining.TrueValue
		Me.Refining.Width
 i
 "Refrigeration
 1}
 = 25
 = True
 ("Refining.NullvValue"),
 System.
 System.
 System.
		Me.Refrigeration.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Refrigeration.FalseValue
		Me.Refrigeration.HeaderText
		Me.Refrigeration.MappingName
		Me.Refrigeration.NullValue
 System.DBNull)
		Me.Refrigeration.ReadOnly
		Me.Refrigeration.TrueValue
		Me.Refrigeration.Width
 f
 "Republic
 §
		Me.Republic.Alignment
 = 25
 = False
 = "Rfr"
 = "Refrigeration"
 = CType (resources.GetObject
 = True
 = True
 ("Refrigeration.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Republic.FalseValue = False
		Me.Republic.HeaderText
 = "Rep"
		Me.Republic.MappingName = "Republic"
		Me.Republic.NullValue
 DBNull)
 |
 = CType (resources.GetObject
		Me.Republic.ReadOnly = True
		Me.Republic.TrueValue
		Me. Republic.Width
 ¥
 ‘Robotics
 = 25
 = True
 ("Republic.NullValue"),
 System.
 36
 v4
 «
 v4
 «
 4
 LH
		Me.Robotics.Alignment
 II
 Extractor\Forml.vb
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Robotics.FalseValue = False
		Me.Robotics.HeaderText
 = "Rob"
		Me.Robotics.MappingName = "Robotics"
		Me.Robotics.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Robotics.ReadOnly = True
		Me.Robotics.TrueValue
		Me.Robotics.Width
 ¥
 "Rocketry
 H
 = 25
		Me.Rocketry.Alignment
 = True
 ("Robotics.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Rocketry.FalseValue = False
		Me.Rocketry.HeaderText
 = "Roc"
		Me.Rocketry.MappingName = "Rocketry"
		Me.Rocketry.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Rocketry.ReadOnly = True
		Me.Rocketry.TrueValue
		Me.Rocketry.Width
 t
 ‘Sanitation
 H
 = 25
 = True
 ("Rocketry.NullValue”),
		Me.Sanitation.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Sanitation.FalsevValue = False
		Me.Sanitation.HeaderText
 = "San"
		Me.Sanitation.MappingName = "Sanitation"
		Me.Sanitation.NullValue
 DBNull)
		Me.Sanitation.ReadOnly
		Me.Sanitation.TrueValue
		Me.Sanitation.Width
 Hi
 'Seafaring
 §
		Me.Seafaring.Alignment
 = CType (resources.GetObject("Sanitation.NullvValue"),
 = True
 = True
 = 25
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Seafaring.FalseValue = False
		Me.Seafaring.HeaderText
 = "Sea"
		Me.Seafaring.MappingName = "Seafaring"
		Me. Seafaring.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Seafaring.ReadOnly = True
		Me.Seafaring.TrueValue
		Me.Seafaring.Width
 'SpaceFlight
 = 25
 = True
 ("Seafaring.NullValue"),
 System.
 System.
 System.
		Me.SpaceFlight.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.SpaceFlight.FalseValue = False
		Me.SpaceFlight.HeaderText
 = "g Fp"
		Me.SpaceFlight.MappingName = "SpaceFlight"
		Me.SpaceFlight.NullValue
 System.DBNull)
		Me.SpaceFlight.ReadOnly
		Me. SpaceFlight.TrueValue
		Me.SpaceFlight.Wwidth
 '
 ‘Stealth
		Me.Stealth.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("SpaceFlight.NullvValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Stealth.FalseValue = False
		Me.Stealth.HeaderText = "Ste"
		Me.Stealth.MappingName = "Stealth"
		Me.Stealth.Nullvalue
 = CType (resources.GetObject("Stealth.Nullvalue"),
 37
 v4
 v4
 System. «
 4
 v4
 System.DBNullw
		Me.Stealth.ReadOnly
		Me.Stealth.TrueValue
		Me.Stealth.Wwidth
 = True
 = True
 = 25
 'SteamEngine
 4
 i
		Me.SteamEngine.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.SteamEngine.FalseValue = False
		Me.SteamEngine.HeaderText
		Me.SteamEngine.MappingName
 = "S E"
 = "SteamEngine"
		Me.SteamEngine.NullValue
 = CType (resources.GetObject
 ("SteamEngine.NullValue"),
 System.DBNull)
		Me.SteamEngine.ReadOnly = True
		Me.SteamEngine.TrueValue = True
		Me.SteamEngine.Width
 "Steel
 ¥
		Me.Steel.Alignment
 = 25
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Steel.FalseValue = False
		Me.Steel.HeaderText = "Ste"
		Me.Steel.MappingName = "Steel"
		Me.Steel.NullValue
 = CType (resources.GetObject
		Me.Steel.ReadOnly = True
		Me.Steel.TrueValue
		Me.Steel.Width
 t
 "Superconductor
 ¥
 = 25
 = True
		Me.Superconductor.Alignment
 ("Steel.NullValue"),
 System.DBNull)
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Superconductor.FalseValue = False
		Me. Superconductor.
		Me. Superconductor
 HeaderText
 Sup
 .MappingName = "Superconductor"
		Me.Superconductor.NullValue
 = CType (resources.GetObject
 System.DBNull)
		Me.Superconductor.ReadOnly
		Me.Superconductor.TrueValue
		Me. Superconductor.Width
 !
 ‘Tactics
 ¥
		Me.Tactics.Alignment
 = True
 = True
 = 25
 ("Superconductor.Nullvalue”),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Tactics.FalseValue = False
		Me.Tactics.HeaderText
 = "Tac"
		Me.Tactics.MappingName = "Tactics"
		Me.Tactics.NullValue
 = CType (resources.GetObject
 ("Tactics.Nullvalue"),
 38
 v4
 ¢
 System.DBNullw
		Me.Tactics.ReadOnly
		Me.Tactics.TrueValue
		Me.Tactics.Width
 = True
 = True
 = 25
 "Theology
		Me.Theology.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Theology.FalseValue
 = False
		Me.Theology.HeaderText
		Me.Theology.MappingName
		Me.Theology.NullValue
 DBNull)
 = "The"
 = "Theology"
 = CType (resources.GetObject
		Me. Theology.ReadOnly = True
		Me.Theology.TrueValue = True
		Me.Theology.Width
 ¢
 'TheoryOfGravity
 i
 = 25
		Me.TheoryOfGravity.Alignment
		Me.TheoryOfGravity.FalseValue
		Me.TheoryOfGravity.HeaderText
		Me.TheoryOfGravity.MappingName = "TheoryOfGravity"
		Me.TheoryOfGravity.NullValue
 ) 4
 System.DBNull)
		Me.TheoryOfGravity.ReadOnly
		Me.TheoryOfGravity.TrueValue
		Me.TheoryOfGravity.Width
 Hi
 "Trade
 f
 = 25
 ("Theology.NullvValue"),
 System.
 = System.Windows.Forms.HorizontalAlignment.Center
 = False
 = "TOG"
 = CType (resources.GetObject
 = True
 = True
 ("TheoryOfGravity.NullValue"
		Me.Trade.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Trade.FalseValue = False
 vs
 ¢
		Me.Trade.HeaderText = "Tra"
		Me.Trade.MappingName = "Trade"
		Me.Trade.NullValue
 = CType (resources.GetObject
		Me.Trade.ReadOnly = True
		Me.Trade.TrueValue = True
		Me.Trade.Width
 "University
 H
 = 25
 ("Trade.Nullvalue"),
 System.DBNull)
		Me.University .Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.University .FalseValue = False
		Me.University
		Me.University
		Me.University
 System.DBNull)
		Me.University
 .HeaderText
 = "Uni"
 .MappingName = "University"
 .NullValue
 = CType (resources.GetObject
 .ReadOnly = True
		Me.University .TrueValue = True
		Me.University .Width = 25
 H
 'WarriorCode
 ¥
		Me.WarriorCode.Alignment
 ("University
 .Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.WarriorCode.FalseValue = False
		Me.WarriorCode.HeaderText
		Me.WarriorCode.MappingName
		Me.WarriorCode.NullValue
 System.DBNull)
 = "War"
 = "WarriorCode™
 = CType (resources.GetObject
		Me.WarriorCode.ReadOnly = True
		Me.WarriorCode.TrueValue
		Me.WarriorCode.Width
 L
 "Wheel
 §
		Me.Wheel.
 = 25
 = True
 ("WarriorCode.Nullvalue"),
 Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Wheel.
		Me.Wheel.
		Me.Wheel.
 FalseValue = False
 HeaderText
 = "Whe"
 MappingName = "Wheel"
		Me.Wheel.NullValue
 = CType (resources.GetObject
		Me.Wheel.ReadOnly = True
		Me.Wheel.TrueValue = True
		Me.Wheel.Width
 t
 "Writing
 4
 = 25
		Me.Writing.Alignment
 ("Wheel.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Writing.FalseValue = False
		Me.Writing.HeaderText
 = "Wri"
		Me.Writing.MappingName = "Writing"
		Me.Writing.NullValue
 = CType (resources.GetObject
 ("Writing.NullvValue"),
 System.DBNull)
 39
 v4
 v4
 System.DBNullw
		Me.Writing.ReadOnly = True
		Me.Writing.TrueValue = True
		Me.Writing.Width
 End Sub
 #End Region
 #Region "Cities"
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 = 25
 CityNumber
 As CIV II Extractor.DataGridColouredTextBoxColumn
 CityName As CIV II Extractor.DataGridColouredTextBoxColumn
 CityHorizCoord
 CityVertCoord
 CityOwnerColour
 CitySize
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II ExXtractor.DataGridColouredTextBoxColumn
 Friend
 Friend
 WithEvents
 WithEvents
 OrigColour
 As CIV _II Extractor.DataGridColouredTextBoxColumn
 WorkingCitySquaresCount
 DataGridColouredTextBoxColumn
 Friend
 WithEvents
 Palace
 As CIV II Extractor.
 RY SF
 As CIV II Extractor.DataGridColouredBoolColumn
 Friend
 Friend
 WithEvents
 WithEvents
 Barracks
 Granary
 As CIV II Extractor.DataGridColouredBoolColumn
 As CIV II Extractor.DataGridColouredBoolColumn
 Friend WithEvents Temple As CIV II Extractor.DataGridColouredBoolColumn
 Friend
 Friend
 WithEvents
 WithEvents
 Marketplace
 Library
 As CIV II Extractor.DataGridColouredBoolColumn
 v
 As CIV_II Extractor.DataGridColouredBoolColumn
C:\Documents and Settings\Jeremy Flowers\My... II Extractor\Forml.vb 40
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 Courthouse AsCIVIIExtractor.DataGridColouredBoolColumn CityWalls AsCIVIT_Extractor. DataGridColouredBoolColumn
 Aqueduct As CIV II Extractor.DataGridColouredBoolColumn
 Bank As CIV II Extractor.DataGridColouredBoolColumn
 Cathedral As CIV II Extractor.DataGridColouredBoolColumn University AsCIVIIExtractor.DataGridColouredBoolColumn
 MassTransit As CIV IT_Extractor.DataGridColouredBoolColumn
 ColosseumAsCIVITExtractor. DataGridColouredBoolColumn
 Factory As CIV II Extractor. DataGridColouredBoolColumn
 ManufacturingPlant As CIV II Extractor.DataGridColouredBoolColumn
 SDIDefense As CIV II Extractor.DataGridColouredBoolColumn
 RecyclingCentre As CIVII Extractor.DataGridColouredBoolColumn
 PowerPlant As CIV II Extractor.DataGridColouredBoolColumn
 HydroPlant As CIV IT Extractor.DataGridColouredBoolColumn
 NuclearPlant As CIV II Extractor.DataGridColouredBoolColumn
 StockExchange As CIVII Extractor.DataGridColouredBoolColumn
 SewerSystem As CIV II Extractor.DataGridColouredBoolColumn
 Supermarket As CIV ITIExtractor.DataGridColouredBoolColumn
 Superhighways As CIV II Extractor.DataGridColouredBoolColumn
 ResearchLab As CIV IIExtractor.DataGridColouredBoolColumn
 SAMMissileBattery As CIV II Extractor.DataGridColouredBoolColumn
 CoastalFortress As CIV II Extractor.DataGridColouredBoolColumn
 SolarPlant As CIVIIExtractor.DataGridColouredBoolColumn
 Harbor As CIV II Extractor.DataGridColouredBoolColumn
 OffshorePlatform AsCIV IIExtractor.DataGridColouredBoolColumn
 Airport As CIV ITExtractor.DataGridColouredBoolColumn
 PoliceStation As CIV II Extractor.DataGridColouredBoolColumn
 PortFacility As CIV _IIExtractor.DataGridColouredBoolColumn
 CityProducing As CIV TE_Extractor.DataGridColouredTextBoxColumn
 NumberofActiveTradeRoutes AsCIVITExtractor.
 DataGridColouredTextBoxColumn
 SuppliedCommodityl
 SuppliedCommodity?2
 SuppliedCommodity3
 DemandedCommodityl
 DemandedCommodity?2
 DemandedCommodity3
 As
 As
 As
 As
 As
 AS
 CIV II Extractor.DataGridColouredTextBoxColumn
 CIV ITIExtractor.DataGridColouredTextBoxColumn
 CIVIIExtractor.DataGridColouredTextBoxColumn
 CIV IT_Extractor. DataGridColouredTextBoxColumn
 CIVIIExtractor.DataGridColouredTextBoxColumn CIV_II Extractor.DataGridColouredTextBoxColumn
 TradedCommodityl As CIV II Extractor.DataGridColouredTextBoxColumn
 TradedCommodity2 As CIVIIExtractor.DataGridColouredTextBoxColumn
 TradedCommodity3 As CIV4 §
 “Extractor. DataGridColouredTextBoxColumn
 TradingCityNumberl As CIV II _Extractor.DataGridColouredTextBoxColumn
 TradingCityNumber2 As CIV_IT“Extractor. DataGridColouredTextBoxColumn
 TradingCityNumber3 As CIV TE“Extractor. DataGridColouredTextBoxColumn
 ElvisCount AsCIVIT_Extractor. DataGridColouredTextBoxColumn
 ScientistCount As“CIV II Extractor. DataGridColouredTextBoxColumn
 TaxCollectorCount As CIV II Extractor.DataGridColouredTextBoxColumn
 FoodInStorage As CIV II Extractor.DataGridColouredTextBoxColumn
 ShieldsConsumedInProduction As CIV II Extractor.
 DataGridColouredTextBoxColumn oT
 TradeFromCitySquares As CIV IT Extractor.DataGridColouredTextBoxColumn
 ScienceForCity As CIVII Extractor.DataGridColouredTextBoxColumn
 TaxForCity As CIV II Extractor.DataGridColouredTextBoxColumn
 TotalTradeForCity As CIV II Extractor.DataGridColouredTextBoxColumn
 FoodFromCitySquares As CIV II Extractor.DataGridColouredTextBoxColumn
 ShieldsFromCity As CIV II Extractor.DataGridColouredTextBoxColumn
 HappyCitizensInCity AS"CIV II Extractor.DataGridColouredTextBoxColumn UnhappyCitizensInCityAsCIVII Extractor.
 DataGridColouredTextBoxColumn
 Private Sub SetupCityColumns (ByRef resources As System.Resources.ResourceManager)
		Me.
 .CityName
		Me.CityHorizCoord
		Me.CityVertCoord =
		Me.CityOwnerColour
 .CitySize
 .OrigColour
		Me.WorkingCitySquaresCount
		Me.
 .Barracks
		Me.
		Me
		Me
		Me
		Me
		Me.Temple
		Me.Marketplace
 CityNumber
 Palace
 Granary =
 New CIV II Extractor.DataGridColouredTextBoxColumn ()
 =NewCIV IIExtractor.DataGridColouredTextBoxColumn ()
 New CIV II Extractor.DataGridColouredTextBoxColumn ()
 New CIV IE_Extractor. DataGridColouredTextBoxColumn ()
 New CIV II Extractor. DataGridColouredTextBoxColumn ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn ()
 New CIV II Extractor.DataGridColouredBoolColumn ()
 New CIV _II Extractor.DataGridColouredBoolColumn ()
 NewCIV ITExtractor.DataGridColouredBoolColumn ()
 ——
 = New CIV II Extractor.DataGridColouredBoolColumn ()
 New CIV II Extractor.DataGridColouredBoolColumn()
 "4
 '4
		Me.Library
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.Courthouse
		Me.CityWalls
		Me.Adqueduct
		Me.Bank
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV TH _Extractor.
 = New CIV II _Extractor.
		Me.Cathedral
 DataGridColouredBoolColumn
 DataGridColouredBoolColumn
 = New CIV ITExtractor.DataGridColouredBoolColumn
		Me.University
		Me.MassTransit
		Me.Colosseum
		Me.Factory
 = New CIV II _Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.
 DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.ManufacturingPlant
		Me.SDIDefense
		Me.RecyclingCentre
		Me. PowerPlant
		Me.HydroPlant
		Me.NuclearPlant
		Me.StockExchange
		Me. SewerSystem
		Me.Supermarket
		Me.Superhighways
		Me.ResearchlLab
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Ewvtractor.DataCridCslouredBooclColumn
 = New CIV IT Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn()
 = New CIV ITI Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.SAMMissileBattery
 ()
 ()
 ()
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.CoastalFortress
		Me.SolarPlant
		Me.Harbor
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV ITI Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.OffshorePlatform
		Me.Airport
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn()
		Me.PoliceStation
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.PortFacility
		Me.CityProducing
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.NumberofActiveTradeRoutes
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me. SuppliedCommodityl
		Me. SuppliedCommodity2
		Me. SuppliedCommodity3
		Me.DemandedCommodityl
		Me. DemandedCommodity2
		Me. DemandedCommodity3
		Me.TradedCommodityl
		Me.TradedCommodityZz
		Me.TradedCommodity3
		Me.TradingCityNumberl
		Me.TradingCityNumberZ
		Me.TradingCityNumber3
		Me.ElvisCount
 = New CIV ITExtractor.
 = New CIV II Extractor.
 DataGridColouredTextBoxColumn
 DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn()
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV IT_Extractor.
		Me.ScientistCount
 DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.TaxCollectorCount
		Me.FoodInStorage
		Me.ShieldsConsumedInProduction
 = New CIV II Extractor.
 DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.TradeFromCitySquares
		Me.ScienceForCity
		Me.TaxForCity
 = New CIV II Extractor.DataGridColouredTextBoxColumn()
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV IT Extractor.DataGridColouredTextBoxColumn
		Me.TotalTradeForCity
		Me.FoodFromCitySquares
		Me.ShieldsFromCity
		Me.HappyCitizensInCity
		Me.UnhappyCitizensInCity
 m CitiesCS
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New System.Windows.Forms.DataGridColumnStyle()
		Me.CityName, Me.CityHorizCoord,
		Me.CityVertCoord,
 {Me.CityNumber,
		Me.CityOwnerColour,
		Me.CitySize, Me.OrigColour, Me.WorkingCitySquaresCount, Me.Palace,
		Me.Barracks,
		Me.Granary,
		Me.Temple, Me.Marketplace,
		Me.Library,
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
		Me.Courthouse,
		Me.CityWalls,
		Me.Aqueduct, Me.Bank, Me.Cathedral,
		Me.University,
		Me.MassTransit, Me.Colosseum, Me.Factory, Me.ManufacturingPlant,
		Me.SDIDefense, Me.RecyclingCentre,
		Me.NuclearPlant,
		Me.StockExchange,
		Me.PowerPlant, Me.HydroPlant,
		Me.SewerSystem,
		Me. Supermarket,
 Lo
		Me.Superhighways, Me.ResearchlLab, Me.SAMMissileBattery, Me.CoastalFortress,
		Me.SolarPlant,
		Me.PortFacility,
		Me.Harbor, Me.OffshorePlatform,
		Me.CityProducing,
		Me.Airport, Me.PoliceStation,
		Me.NumberofActiveTradeRoutes,
		Me. SuppliedCommodityl,
		Me.SuppliedCommodity2,
		Me.SuppliedCommodity3,
		Me.DemandedCommodityl, Me.DemandedCommodity2, Me.DemandedCommodity3,
		Me.TradedCommodityl,
		Me.TradedCommodity2,
		Me.TradingCityNumberl,
		Me.TradedCommodity3,
		Me.TradingCityNumberZ, Me.TradingCityNumber3,
 41
 ()
 ()
 «
		Me.ElvisCount, Me.ScientistCount,
		Me.TaxCollectorCount,
		Me.FoodInStorage,
		Me.ScienceForCity,
		Me.ShieldsConsumedInProduction,
		Me.TaxForCity, Me.TotalTradeForCity,
		Me.TradeFromCitySquares,
 _
		Me.FoodFromCitySquares, Me.ShieldsFromCity, Me.HappyCitizensInCity,
		Me.UnhappyCitizensInCity}
 §
 'CityNumber
 §
		Me.CityNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.CityNumber.Format
 = ""
		Me.CityNumber.FormatInfo = Nothing
		Me.CityNumber.HeaderText = "City #"
		Me.CityNumber.MappingName
 = "CityNumber"
		Me.CityNumber.ReadOnly = True
		Me.CityNumber.Width
 'CityName
 1
		Me.CityName.Format
 = 45
 = ""
		Me.CityName.FormatInfo = Nothing
		Me.CityName.HeaderText
 = "City Name"
		Me.CityName.MappingName = "CityName"
		Me.CityName.ReadOnly = True
		Me.CityName.Width
 'CityVertCoord
 i
 = 101
		Me.CityVertCoord.Alignment
		Me.CityVertCoord.Format
		Me.CityVertCoord.FormatInfo
		Me.CityVertCoord.HeaderText
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "Vert Coord"
		Me.CityVertCoord.MappingName = "VertCoord”
		Me.CityVertCoord.ReadOnly = True
		Me.CityVertCoord.Width
 ¥
 'CityHorizCoord
 ¥
 = 70
		Me.CityHorizCoord.Alignment
		Me.CityHorizCoord.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.CityHorizCoord.FormatInfo
		Me.CityHorizCoord.HeaderText
 = Nothing
 = "Horiz Coord"
		Me.CityHorizCoord.MappingName = "HorizCoord"
		Me.CityHorizCoord.ReadOnly = True
		Me.CityHorizCoord.Width
 |
 "CityOwnerColour
 4
		Me.CityOwnerColour.Format
 = 70
 = ""
		Me.CityOwnerColour.FormatInfo
		Me.CityOwnerColour.HeaderText
		Me.CityOwnerColour.MappingName
 = Nothing
 = "Owner Colour"
 = "OwnerColour"
		Me.CityOwnerColour.ReadOnly = True
		Me.CityOwnerColour.wWidth
 H
 'CitySize
		Me.CitySize.Alignment
		Me.CitySize.Format
		Me.CitySize.FormatInfo
 = 75
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
		Me.CitySize.HeaderText = "Size"
		Me.CitySize.MappingName = "CitySize"
		Me.CitySize.ReadOnly
		Me.CitySize.Width
 = True
 = 45
 {
 "OrigColour
		Me.OrigColour.Format
 = ""
		Me.OrigColour.FormatInfo
		Me.OrigColour.HeaderText
 = Nothing
 = "Orig Colour"
		Me.OrigColour.MappingName = "OrigColour"
		Me.OrigColour.ReadOnly = True
		Me.OrigColour.Width
 42
 = 65
 1
		Me.WorkingCitySquaresCount.
 Right
		Me.WorkingCitySquaresCount.Format
 Alignment = System.Windows.Forms.HorizontalAlignment.
 = ""
		Me.WorkingCitySquaresCount.FormatInfo
		Me.WorkingCitySquaresCount.HeaderText
 = Nothing
 = "#Workers"
		Me.WorkingCitySquaresCount.MappingName = "WorkingCitySquaresCount"
		Me.WorkingCitySquaresCount.ReadOnly
		Me.WorkingCitySquaresCount.Width
 = True
 = 55
 f
 "Palace
 ¥
		Me.Palace.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Palace.FalseValue = False
		Me.Palace.HeaderText = "Pal"
		Me.Palace.MappingName = "Palace"
		Me.Palace.NullValue
 = CType (resources.GetObject
		Me.Palace.ReadOnly = True
		Me.Palace.TrueValue
		Me.Palace.Width
 ¥
 ‘Barracks
		Me.Barracks
		Me.Barracks
		Me.Barracks
		Me.Barracks
		Me.Barracks
 DBNull)
		Me.Barracks
 = 25
 = True
 ("Palace.NullValue"),
 Alignment = System.Windows.Forms.HorizontalAlignment.Center
 .FalseValue = False
 iHeaderText = "Bar"
 .MappingName = "Barracks"
 .NullValue = CType(resources.GetObject
 .ReadOnly = True
		Me.Barracks
		Me.Barracks
		Me.Granary.
 .TrueValue = True
 Width
 = 25
 ("Barracks.NullValue"),
 Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Granary.
		Me.Granary.
		Me.Granary.
		Me.Granary.
 FalseValue = False
 HeaderText = "Gra"
 MappingName = "Granary"
 NullvValue
 = CType (resources.GetObject
 ("Granary.NullvValue"),
 System.DBNull)
 System.
 43
 System.DBNullw
		Me.Granary.
		Me.Granary.
		Me.Granary.
 'Temple
 ReadOnly = True
 TrueValue = True
 width
 = 25
 .
		Me.Temple.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Temple.FalseValue = False
		Me.Temple.HeaderText
 = "Tmp"
		Me.Temple.MappingName = "Temple"
		Me.Temple.NullValue
 = CType (resources.GetObject
		Me.Temple.ReadOnly = True
		Me.Temple.T
		Me.Temple.Width
 f
 '‘Marketplac
 ¥
		Me.Marketpl
 rueValue = True
 = 25
 €
 ("Temple.NullValue"),
 System.DBNull)
 ace.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Marketpl
		Me.Marketpl
		Me.Marketpl
		Me.Marketpl
 System.DBNull)
		Me.Marketpl
		Me.Marketpl
		Me.Marketpl
 ‘Library
 ¥
 ace.FalseValue = False
 ace .HeaderText
 = "Mkt"
 ace.MappingName = "Marketplace"
 ace.NullValue
 = CType (resources.GetObject
 ace.ReadOnly = True
 ace.TrueValue = True
 ace.Width
 = 25
 ("Marketplace.NullValue"),
 Alignment = System.Windows.Forms.HorizontalAlignment.Center
 v4
 "4
 v4
		Me.Library.
		Me.Library.FalseValue = False
		Me.Library.HeaderText
 = "Lib"
		Me.Library.MappingName = "Library"
		Me.Library.NullValue
 = CType (resources.GetObject
 ("Library.NullValue"),
 44
 System.DBNullw
		Me.Library.ReadOnly = True
		Me.Library.TrueValue
		Me.Library.Width
 24
 "Courthouse
 H
 = 25
 = True
		Me.Courthouse.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Courthouse.FalseValue = False
		Me.Courthouse.HeaderText = "Crt"
		Me.Courthouse.MappingName = "Courthouse"
		Me.Courthouse.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Courthouse.ReadOnly = True
		Me.Courthouse.TrueValue = True
		Me.Courthouse.Width
 §
 'CityWalls
 {
 = 25
		Me.CityWalls.Alignment
 ("Courthouse.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.CityWalls.FalseValue = False
		Me.CityWalls.HeaderText
 = "Wal"
		Me.CityWalls.MappingName = "CityWalls"
		Me.CityWalls.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.CityWalls.ReadOnly = True
		Me.CityWalls.TrueValue
		Me.CityWalls.Width
 "Aqueduct
 = 25
 = True
 ("CityWalls.Nullvalue™),
		Me.Aqueduct.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Aqueduct.FalseValue = False
		Me.Aqueduct.HeaderText
		Me.Aqueduct.MappingName
		Me.Aqueduct.NullValue
 DBNull)
 = "Aqu"
 = "Aqueduct"
 = CType (resources.GetObject
		Me.Aqueduct.ReadOnly = True
		Me.Aqueduct.TrueValue = True
		Me.Aqueduct.Width
 f
 "Bank
 §
 = 25
 ("Aqueduct .NullValue"),
		Me.Bank.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Bank.FalseValue = False
		Me.Bank.HeaderText
 = "Bnk"
		Me.Bank.MappingName = "Bank"
		Me.Bank.NullvValue = CType (resources.GetObject
 ("Bank.NullValue"),
 System. ¢
 System.
 System.
 System.DBNull)
		Me.Bank.ReadOnly = True
		Me.Bank.TrueValue = True
		Me.Bank.Width
 t
 "Cathedral
 H
 = 25
		Me.Cathedral.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Cathedral.FalseValue = False
		Me.Cathedral.HeaderText
 = "Cat"
		Me.Cathedral.MappingName = "Cathedral"
		Me.Cathedral.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Cathedral.ReadOnly = True
		Me.Cathedral.TrueValue
		Me.Cathedral.Width
 "University
 = 25
 :
		Me.University.Alignment
 = True
		Me.University.FalseValue = False
 ("Cathedral.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
 = "Uni"
 System.
 «
 ¥
 v4
		Me.University.HeaderText
		Me.University.MappingName = "University"
		Me.University.NullvValue
 = CType (resources.GetObject
 ("University.NullValue"™),
 DBNull)
		Me.University.ReadOnly
		Me.University.TrueValue
		Me.University.Width
 f
 'MassTransit
 ¥
 = True
 = True
 = 25
		Me.MassTransit.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.MassTransit.FalseValue = False
		Me.MassTransit.HeaderText
 "M T"
		Me.MassTransit.MappingName = "MassTransit"
		Me.MassTransit.NullValue
 System.DBNull)
		Me.MassTransit.ReadOnly
		Me.MassTransit.TrueValue
		Me.MassTransit.Width
 §
 "Colosseum
 f
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("MaggTranegit.NullValue},
		Me.Colosseum.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Colosseum.FalseValue = False
		Me.Colosseum.HeaderText = "Col"
		Me.Colosseum.MappingName = "Colosseum"
		Me.Colosseum.NullvValue
 DBNull)
 = CType (resources.GetObject
		Me.Colosseum.ReadOnly = True
		Me.Colosseum.TrueValue = True
		Me.Colosseum.Width
 ‘Factory
 ¥
		Me.Factory.Alignment
 = 25
 ("Colosseum.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Factory.FalseValue = False
 l
		Me.Factory.HeaderText
 = "Fac"
		Me.Factory.MappingName = "Factory"
		Me.Factory.NullValue
 = CType (resources.GetObject
		Me.Factory.ReadOnly = True
		Me.Factory.TrueValue
		Me.Factory.Width
 f
 ‘*ManufacturingPlant
 ¥
 = 25
 = True
		Me.ManufacturingPlant.Alignment
 ("Factory.NullValue"),
 System.
 System.DBNull
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.ManufacturingPlant.FalseValue = False
		Me.ManufacturingPlant.HeaderText
		Me.ManufacturingPlant.MappingName
		Me.ManufacturingPlant.NullValue
 NullvValue"), System.DBNull)
		Me.ManufacturingPlant.ReadOnly
 = "MPL"
 = "ManufacturingPlant"
 = CType (resources.GetObject
 = True
		Me.ManufacturingPlant.TrueValue
		Me.ManufacturingPlant.Width
 H
 'SDIDefense
 f
 = True
 = 25
 ("ManufacturingPlant.
		Me.SDIDefense.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.SDIDefense.FalseValue = False
		Me.SDIDefense.HeaderText
		Me.SDIDefense.MappingName
		Me.S3DIDefense.NullValue
 DBNull)
 = "SDI"
 = "SDIDefense"
 = CType (resources.GetObject
		Me.SDIDefense.ReadOnly = True
		Me.SDIDefense.TrueValue = True
		Me.SDIDefense.Width
 ¥
 "RecyclingCentre
 = 25
		Me.RecyclingCentre.Alignment
 ("SDIDefense.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.RecyclingCentre.FalseValue = False
		Me.RecyclingCentre.HeaderText
 = "Rcy"
 45
 System. «
 v4
 «
 «
 "4
 System. wv
		Me.RecyclingCentre.MappingName = "RecyclingCentre™
		Me.RecyclingCentre.NullValue
 ),
 System.DBNull)
		Me.RecyclingCentre.ReadOnly
 = CType (resources.GetObject
 = True
		Me.RecyclingCentre.TrueValue
		Me.RecyclingCentre.Width
 = 25
 = True
 ("RecyclingCentre.NullvValue”
		Me.PowerPlant.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.PowerPlant.FalseValue = False
 |
		Me.PowerPlant.HeaderText
 = "Pow"
		Me.PowerPlant.MappingName = "PowerPlant"
		Me.PowerPlant.NullValue
 = CType (resources.GetObject
 ("PowerPlant.NullValue"),
 46
 ¢
 System. «
 DBNull)
		Me. PowerPlant.ReadOnly = True
		Me.PowerPlant.TrueValue
		Me.PowerPlant.Width
 = 25
		Me.HydroPlant.Alignment
 = True
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.HydroPlant.FalseValue = False
		Me.HydroPlant.HeaderText
 = "Hyd"
		Me.HydroPlant.MappingName = "HydroPlant"
		Me.HydroPlant.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.HydroPlant.ReadOnly = True
		Me.HydroPlant.TrueValue
		Me.HydroPlant .Width = 25
 <
 |
 os
 § wep
 FR
 ome bom emmanTF] emmen ree
 LS
 CA ve A de CALL Ge
 § fous ped FF Be
 ~~ 8 § 3%
		Me.NuclearPlant.Alignment
 = True
 ("HydroPlant.NullvValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.NuclearPlant.FalseValue = False
		Me.NuclearPlant.HeaderText
 = "Nuc"
		Me.NuclearPlant.MappingName = "NuclearPlant"
		Me.NuclearPlant.NullValue
 System.DBNull)
		Me.NuclearPlant.ReadOnly
		Me.NuclearPlant.TrueValue
		Me.NuclearPlant.Width
 ry.
 fF Ng
 5
 ¢ [oo IY uy
 wr lean
|
 lldil LW
 i.
 Sw
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("NuclearPlant.NullvValue"},
		Me.StockExchange.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.StockExchange.FalseValue = False
		Me.StockExchange.HeaderText
 = "Stk"
		Me.StockExchange.MappingName = "StockExchange"
		Me.StockExchange.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.StockExchange.ReadOnly = True
		Me.StockExchange.TrueValue = True
		Me.StockExchange.Width
 f
 CF yw yom am Cf
 Seweray
 L
 1]
 >
 IF)
 = 25
 ("StockExchange.NullValue"),
		Me.SewerSystem.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.SewerSystem.FalseValue = False
		Me.SewerSystem.HeaderText
		Me.SewerSystem.MappingName
		Me.SewerSystem.NullValue
 System.DBNull)
 = "Sew"
 = "SewerSystem"
 = CType (resources.GetObject
		Me.SewerSystem.ReadOnly = True
		Me.SewerSystem.TrueValue = True
		Me.SewerSystem.Width
 }
 ¥
 bod ad CVPR SERS RR=
 Hr
 ud
 Lal ral
		Me. Supermarket.Alignment
 = 25
 ("SewerSystem.Nullvalue”),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Supermarket.FalseValue = False
		Me.Supermarket.HeaderText
		Me. Supermarket.MappingName
		Me.Supermarket.NullValue
 = "Smk"
 = "Supermarket"
 = CType (resources.GetObject
 System. «
 4
 v4
 "4
 ("Supermarket.Nullvalue"),
 "4
 System.DBNull)
		Me. Supermarket.ReadOnly
		Me. Supermarket.TrueValue
		Me. Supermarket.wWidth
 f
 'Superhighways
		Me.Superhighways.Alignment
 = True
 = True
 = 25
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Superhighways.FalseValue = False
		Me.Superhighways.HeaderText
 = "Shw"
		Me.Superhighways.MappingName = "Superhighways"
		Me. Superhighways.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.Superhighways.ReadOnly = True
		Me.Superhighways.TrueValue
		Me. Superhighways.Width
 §
 ‘ResearchLab
 H
 = 25
 = True
 ("Superhighways.NullValue"),
		Me.ResearchlLab.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Researchlab.FalseValue = False
		Me.ResearchlLab.HeaderText = "Res"
		Me.ResearchLab.MappingName = "ResearchLab"
		Me.ResearchlLab.NullValue
 System.DBNull)
 = CType (resources.GetObject
		Me.ResearchlLab.ReadOnly = True
		Me.ResearchlLab.TrueValue = True
		Me.ResearchLab.Width
 'SAMMissileBattery
 = 25
		Me.SAMMissileBattery.Alignment
 ("ResearchLab.NullvValue"},
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.SAMMissileBattery.FalseValue = False
		Me.SAMMissileBattery.HeaderText
 = "SAM"
		Me.SAMMissileBattery.MappingName = "SAMMissileBattery"
		Me.SAMMissileBattery.NullValue
 NullvValue"), System.DBNull)
 = CType (resources.GetObject
		Me.SAMMissileBattery.ReadOnly = True
		Me.SAMMissileBattery.TrueValue
		Me.SAMMissileBattery.Width
 ¥
 '*CoastalFortress
 4
		Me.CoastalFortress.Alignment
 = 25
 = True
 ("SAMMissileBattery.
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.CoastalFortress.FalseValue
		Me.CoastalFortress.HeaderText
		Me.CoastalFortress.MappingName
		Me.CoastalFortress.NullValue
 ),
 System.DBNull)
		Me.CoastalFortress.ReadOnly
		Me.CoastalFortress.TrueValue
		Me.CoastalFortress.Width
 'SolarPlant
		Me.SolarPlant.Alignment
 = 25
 = False
 = "CsFEF"
 = "CoastalFortress™
 = CType (resources.GetObject
 = True
 = True
 ("CoastalFortress.NullvValue"
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.SolarPlant.FalseValue = False
		Me.SolarPlant.HeaderText = "Slr"
		Me.SolarPlant.MappingName = "SolarPlant"
		Me.SolarPlant.NullValue
 DBNull)
		Me.SolarPlant.ReadOnly
		Me.SolarPlant.TrueValue
		Me.SolarPlant.Width
 1
 "Harbor
 ¥
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("SolarPlant.NullValue"),
		Me.Harbor.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Harbor.FalseValue
		Me.Harbor.HeaderText
		Me.Harbor.MappingName
		Me.Harbor.NullValue
 False
 = "Har"
 = "Harbor"
 = CType (resources.GetObject
 ("Harbor.NullValue"),
 System.DBNull)
 47
 System.
 v4
 v4
 v4
 «¢
		Me.Harbor.ReadOnly = True
		Me.Harbor.TrueValue = True
		Me.Harbor.Width
 §
 "Of fshorePlatform
 ¢
 = 25
		Me.OffshorePlatform.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.OffshorePlatform.FalseValue = False
		Me.OffshorePlatform.HeaderText
 = "Off"
		Me.OffshorePlatform.MappingName = "OffshorePlatform”
		Me.OffshorePlatform.NullValue
 NullValue"),
 System.DBNull)
		Me.OffshorePlatform.ReadOnly
		Me.OffshorePlatform.TrueValue
		Me.OffshorePlatform.Width
 "Airport
 f
		Me.Alrport.Alignment
 = CType (resources.GetObject
 = True
 = True
 = 25
 ("OffshorePlatform.
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Alrport.FalseValue = False
		Me.Airport.HeaderText
 = "Air"
		Me.Airport.MappingName = "Airport"
		Me.Airport.NullValue
 = CType (resources.GetObject
 ("Airport.NullValue"),
 48
 System.DBNullw«
		Me.Alrport.ReadOnly
		Me.Alrport.TrueValue
		Me.Airport.Width
 = True
 = True
 = 25
 f
 'PoliceStation
 H
		Me.PoliceStation.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.PoliceStation.FalseValue = False
		Me.PoliceStation.HeaderText
		Me.PoliceStation.MappingName
		Me.PoliceStation.NullValue
 System.DBNull)
		Me.PoliceStation.ReadOnly
		Me.PoliceStation.TrueValue
		Me.PoliceStation.Width
 'PortFacility
 T
		Me.PortFacility.Alignment
 = 235
 = "Pol"
 = "PoliceStation"
 = CType (resources.GetObject
 = True
 = True
 ("PoliceStation.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.PortFacility.FalseValue
		Me.PortFacility.HeaderText
		Me.PortFacility.MappingName
		Me.PortFacility.NullValue
 System.DBNull)
		Me.PortFacility.ReadOnly
		Me.PortFacility.TrueValue
		Me.PortFacility.Width
 1
 'CityProducing
		Me.CityProducing.Format
 = 25
		Me.CityProducing.FormatInfo
		Me.CityProducing.HeaderText
 = False
 = "Por"
 = "PortFacility"™
 = CType (resources.GetObject
 = True
 = True
 = ""
 = Nothing
 = "Producing"
		Me.CityProducing.MappingName = "CityProducing”
		Me.CityProducing.ReadOnly = True
		Me.CityProducing.Width
 4
 '‘NumberofActiveTradeRoutes
 = 75
		Me.NumberofActiveTradeRoutes.Alignment
 Right
		Me. NumberofActiveTradeRoutes.Format
		Me.NumberofActiveTradeRoutes.FormatInfo
		Me. NumberofActiveTradeRoutes.HeaderText
		Me.NumberofActiveTradeRoutes.MappingName
		Me.NumberofActiveTradeRoutes.ReadOnly
		Me. NumberofActiveTradeRoutes.Width
 §
 'SuppliedCommodityl
 ¥
 ("PortFacility.Nullvalue"),
 = System.Windows.Forms.HorizontalAlignment.
 = ""
 = Nothing
 = "Active Trade"
 = "NumberOfActiveTradeRoutes"”
 = True
 = 75
 "4
 v4
 "4
 «
		Me.SuppliedCommodityl.Format
		Me. SuppliedCommodityl.FormatInfo
		Me. SuppliedCommodityl.HeaderText
 = ""
 = Nothing
 = "Supply
		Me. SuppliedCommodityl.MappingName
 Comm 1"
 = "SuppliedCommodityl"
		Me.SuppliedCommodityl.ReadOnly = True
		Me.SuppliedCommodityl.Width
 'SuppliedCommodityZ
		Me. SuppliedCommodity2.Format
 = 95
 = ""
		Me.SuppliedCommodity2.FormatInfo
		Me. SuppliedCommodity2.HeaderText
		Me. SuppliedCommodity2.MappingName
 = Nothing
 = "Supply
 Comm 2"
 = "SuppliedCommodity2"
		Me. SuppliedCommodity2.ReadOnly
		Me.SuppliedCommodity2.Width
 "SuppliedCommodity3
		Me.SuppliedCommodity3.Format
		Me.SuppliedCommodity3.FormatInfo
		Me. SuppliedCommodity3.HeaderText
		Me.SuppliedCommodity3.MappingName
		Me. SuppliedCommodity3.ReadOnly
		Me. SuppliedCommodity3.Width
 §
 'DemandedCommodityl
 f
		Me.DemandedCommodityl.Format
 = True
 = 95
 = ""
 =
 = 95
 = ""
 =
 =
 Nothing
 "Supply
 =
 True
		Me. DemandedCommodityl.FormatInfo = Nothing
		Me.DemandedCommodityl
 .HeaderText
 Comm 3"
 "SuppliedCommodity3"
 = "Demand Comm 1"
		Me. DemandedCommodityl.MappingName = "DemandedCommodityl"
		Me.DemandedCommodityl. ReadOnly = True
		Me.DemandedCommodityl.Width
 H
 'DemandedCommodityZ
 3
		Me.DemandedCommodity2.Format
 = 95
 = ""
		Me. DemandedCommodity2.FormatInfo = Nothing
		Me.DemandedCommodity2.HeaderText
 = "Demand Comm 2"
		Me. DemandedCommodity?2.MappingName = "DemandedCommodity2"
		Me. DemandedCommodity2.ReadOnly = True
		Me. DemandedCommodity2.Width
 ¥
 DemandedCommodity3
 H
		Me.DemandedCommodity3.Format
 = 95
 = ""
		Me. DemandedCommodity3.FormatInfo = Nothing
		Me. DemandedCommodity3.HeaderText
 = "Demand Comm 3"
		Me.DemandedCommodity3.MappingName = "DemandedCommodity3"
		Me.DemandedCommodity3.ReadOnly
 = True
		Me.DemandedCommodity3.Width
 Li
 ‘TradedCommodityl
 f
		Me.TradedCommodityl.Format
 = 95
 = ""
		Me.TradedCommodityl.FormatInfo
		Me.TradedCommodityl.HeaderText
		Me. TradedCommodityl.MappingName
 = Nothing
 = "Traded Comm 1"
 = "TradedCommodityl"”
		Me.TradedCommodityl.ReadOnly = True
		Me.TradedCommodityl.Width
 t
 :
 'TradedCommodity?Z
		Me.TradedCommodity2.Format
		Me.TradedCommodity2.FormatInfo
		Me.TradedCommodity2.HeaderText
 = 95
 = ""
 =
 =
		Me. TradedCommodity?2 .MappingName =
 Nothing
 "Traded
 Comm 2"
 "TradedCommodity2"
		Me. TradedCommodity2.ReadOnly
		Me.TradedCommodity2.Width
 'TradedCommodity3
 =
 = 95
 = ""
 True
 49
		Me.TradedCommodity3.Format
		Me.TradedCommodity3.FormatInfo = Nothing
		Me.TradedCommodity3.HeaderText
		Me.TradedCommodity3.MappingName
 = "Traded
 Comm 3"
 = "TradedCommodity3"
		Me.TradedCommodity3.ReadOnly = True
		Me.TradedCommodity3.Width
 L
 '"TradingCityNumberl
 H
 = 95
		Me.TradingCityNumberl.Alignment
		Me.TradingCityNumberl.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.TradingCityNumberl.FormatInfo
		Me.TradingCityNumberl.HeaderText
 = Nothing
 = "Trade City 1"
		Me.TradingCityNumberl.MappingName = "TradingCityNumberl"
		Me.TradingCityNumberl.ReadOnly = True
		Me.TradingCityNumberl.Width
 H
 "TradingCityNumber?2
 4
 = 70
		Me.TradingCityNumber2.Alignment
		Me.TradingCityNumber2.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.TradingCityNumber2.FormatInfo
		Me.TradingCityNumber2.HeaderText
 = Nothing
 = "Trade City 2"
		Me.TradingCityNumber2.MappingName = "TradingCityNumber2"
		Me.TradingCityNumber2.ReadOnly = True
		Me.TradingCityNumber2.Width
 ¥
 'TradingCityNumber3
 4
 = 70
		Me.TradingCityNumber3.Alignment
		Me.TradingCityNumber3.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.TradingCityNumber3.FormatInfo
		Me.TradingCityNumber3.HeaderText
 = Nothing
 = "Trade City 3"
		Me.TradingCityNumber3.MappingName = "TradingCityNumber3"
		Me.TradingCityNumber3.ReadOnly = True
		Me.TradingCityNumber3.Width
 t
 'ElvisCount
 i
		Me.ElvisCount.Alignment
		Me.FElvisCount.Format
		Me.ElvisCount.FormatInfo
 = 70
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
		Me.ElvisCount.HeaderText = "# of Elvis"
		Me.ElvisCount.MappingName = "ElvisCount”
		Me.ElvisCount.ReadOnly
		Me.ElvisCount.Width
 = True
 = 50
 "ScientistCount
 4)
		Me.ScientistCount.Alignment
		Me.ScientistCount.Format
		Me.ScientistCount.FormatInfo
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
		Me.ScientistCount.HeaderText = "# of Scientists”
		Me.ScientistCount.MappingName = "ScientistCount"”
		Me. ScientistCount.ReadOnly
		Me.ScientistCount.wWidth
 §
 '"PTaxCollectorCount
 H
 = True
 = 75
		Me.TaxCollectorCount.Alignment
		Me.TaxCollectorCount.Format
		Me.TaxCollectorCount.FormatInfo
		Me.TaxCollectorCount.HeaderText
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "# Of Tax Collectors”
		Me.TaxCollectorCount.MappingName = "TaxCollectorCount”
		Me.TaxCollectorCount.ReadOnly
		Me.TaxCollectorCount.Width
 'FoodInStorage
		Me.FoodInStorage.Alignment
		Me.FoodInStorage.Format
 = ""
		Me.FoodInStorage.FormatInfo
		Me.FoodInStorage.HeaderText
 = True
 = 110
 = System.Windows.Forms.HorizontalAlignment.Right
 Nothing
 5
 = "Stored Food"
		Me.FoodInStorage.MappingName = "FoodInStorage"
		Me.FoodInStorage.ReadOnly
		Me.FoodInStorage.Width
 = True
 = 75
 tf
 'ShieldsConsumedInProduction
		Me.ShieldsConsumedInProduction.Alignment = System.Windows.Forms.HorizontalAlignment.
 Right
		Me.ShieldsConsumedInProduction.Format
		Me.ShieldsConsumedInProduction.FormatInfo
		Me.ShieldsConsumedInProduction.HeaderText
 = ""
		Me.ShieldsConsumedInProduction.MappingName
		Me.ShieldsConsumedInProduction.ReadOnly
 = Nothing
 = "Sheilds
 Consumed"
 = "ShieldsConsumedInProduction"
 = True
 Ma .ShiecldeConcumedInDroduction.
 f
 '"TradeFromCitySquares
 f
		Me.TradeFromCitySquares.Alignment
		Me.TradeFromCitySquares.Format
		Me.TradeFromCitySquares.FormatInfo
 Width = 105
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.TradeFromCitySquares.HeaderText
 = Nothing
 = "Trade from City"
		Me.TradeFromCitySquares.MappingName = "TradeFromCitySquares"
		Me.TradeFromCitySquares.ReadOnly
		Me.TradeFromCitySquares.Width
 = True
 = 85
 'SeienceForCity
 t
		Me.ScienceForCity.Alignment
		Me.ScienceForCity.Format
		Me.ScienceForCity.FormatInfo
		Me.ScienceForCity.HeaderText
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "Science”
		Me.ScienceForCity.MappingName = "ScienceForCity"
		Me.ScienceForCity.ReadOnly
		Me.ScienceForCity.Width
 = True
 = 55
 "TaxForCity
		Me.TaxForCity.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.TaxForCity.Format
 = ""
		Me.TaxForCity.FormatInfo
		Me.TaxForCity.HeaderText
 = Nothing
 = "Tax"
		Me.TaxForCity.MappingName = "TaxForCity"
		Me.TaxForCity.ReadOnly = True
		Me.TaxForCity.Width
 1
 'TotalTradeForCity
 H
 = 55
		Me.TotalTradeForCity.Alignment
		Me.TotalTradeForCity.Format
		Me.TotalTradeForCity.FormatInfo
		Me.TotalTradeForCity.HeaderText
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "Total City Trade"
		Me.TotalTradeForCity.MappingName = "TotalTradeForCity"
		Me.TotalTradeForCity.ReadOnly
		Me.TotalTradeForCity.Width
 t
 'FoodFromCitySquares
 H
		Me.FoodFromCitySquares.Alignment
		Me.FoodFromCitySquares.Format
 = True
 = 90
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.FoodFromCitySquares.FormatInfo
		Me.FoodFromCitySquares.HeaderText
 = Nothing
 = "Food from City"
		Me.FoodFromCitySquares.MappingName = "FoodFromCitySquares”
		Me. FoodFromCitySquares.ReadOnly = True
		Me.FoodFromCitySquares.Width
 §
 "ShieldsFromCity
 L
		Me.ShieldsFromCity.Alignment
		Me.ShieldsFromCity.Format
		Me.ShieldsFromCity.FormatInfo
		Me.ShieldsFromCity.HeaderText
 = 85
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "City Shields"
		Me.ShieldsFromCity.MappingName = "ShieldsFromCity"
		Me.ShieldsFromCity.ReadOnly = True
 51
 «
		Me.ShieldsFromCity.Width
 'HappyCitizensInCity
 = 75
		Me.HappyCitizensInCity.Alignment
 = System.Windows.Forms.HorizontalAlignment.Right
		Me.HappyCitizensInCity.Format
		Me
		Me.
		Me.
		Me
		Me.
 ?
 .HappyCitizensInCity.FormatInfo
 HappyCitizensInCity.HeaderText
 = ""
 = Nothing
 = "Happy Count”
 HappyCitizensInCity.MappingName = "HappyCitizensInCity"
 .HappyCitizensInCity.ReadOnly
 HappyCitizensInCity.Width
 'UnhappyCitizensInCity
 §
		Me
		Me
		Me
 .UnhappyCitizensInCity.Alignment
 .UnhappyCitizensInCity.Format
 .UnhappyCitizensInCity.FormatInfo
 UnhappyCitizensInCity.HeaderText
 = True
 = 75
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me. .UnhappyCitizensInCity.MappingName
		Me
		Me
		Me
 .UnhappyCitizensInCity.ReadOnly
 = Nothing
 = "Unhappy Count”
 = "UnhappyCitizensInCity"
 = True
 .UnhappyCitizensInCity.Width
 End Sub
 #End Region
 #Region "Units"
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 UnitNumber
 UnitHorizCoord
 UnitVertCoord
 = 85
 As CIV ITExtractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 WithEvents Veteran As CIV II Extractor.DataGridColouredBoolColumn
 WithEvents
 UnitType
 As CIV HE_Extractor.DataGridColouredTextBoxColumn
 WithEvents UnitOwnerColour AsTerV IT Extractor.DataGridColouredTextBoxColumn
 WithEvents HitPoints As CIV_II Extractor. DataGridColouredTextBoxColumn
 WithEvents
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 UnitCommodity
 UnitOrders
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 UnitHorizGotoCoords
 UnitVertGotoCoords
 HomeCityNumber
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV TITExtractor.DataGridColouredTextBoxColumn
 WithEvents UnitAbove As CIV II _Extractor. DataGridColouredTextBoxColumn
 WithEvents
 Friend
 WithEvents
 UnitBelow
 As CIV _II _Extractor.
 UnitHomeC1ityName
 DataGridColouredTextBoxColumn
 As CIV IT Extractor.DataGridColouredTextBoxColumn
 Friend
 Friend
 Friend
 Friend
 Private
 WithEvents
 UnitLocation
 As CIV II Extractor.DataGridColouredTextBoxColumn
 WithEvents UnitNationNumber As CIV IT Extractor.DataGridColouredTextBoxColumn
 WithEvents UnitNation As CIV II _Extractor. DataGridColouredTextBoxColumn
 WithEvents UnitNear As CIV ITExtractor.
 DataGridColouredTextBoxColumn
 Sub SetupUnitColumns (ByRef resources As System.Resources.ResourceManhager)
		Me
		Me
		Me
		Me
		Me
 .UnitNumber
 .UnitHorizCoord
 .UnitVertCoord
 .Veteran
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV IT Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV IT Extractor.DataGridColouredBoolColumn
 .UnitType
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me. UnitownerColour
		Me
		Me
 .HitPoints
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 .UnitCommodity
		Me. UnitOrders
		Me
		Me
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV_IT Extractor.DataGridColouredTextBoxColumn
 .UnitHorizGotoCoords
 ()
 ()
 ()
 ()
 ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn-.UnitVertGotoCoords-HomeCityNumber
 .UnitAbove
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV IT Extractor.DataGridColouredTextBoxColumn
 .UnitBelow
 = New CE gk ¥
 _Extractor.
 .UnitHomeCityName
 DataGridColouredTextBoxColumn
 = New CIV. ITExtractor.DataGridColouredTextBoxColumn
 .Unitlocation
 .UnitNationNumber-UnitNation
 = New CIV IT Extractor.
 DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV_II Extractor.DataGridColouredTextBoxColumn
 .UnitNear
 mUnitsCS
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New System.Windows.Forms.DataGridColumnStyle()
 ()
 ()
 ()
 ()
 {Me.UnitNumber,
		Me.UnitHorizCoord, Me.UnitVertCoord, Me.Veteran, Me.UnitType, _
		Me.UnitOwnerColour,
		Me.HitPoints,
		Me.UnitCommodity,
 ()
 ()
 ()
 ()
 ()
		Me.UnitOrders,
		Me.UnitHorizGotoCoords,
		Me.UnitAbove,
		Me.UnitVertGotoCoords,
		Me.UnitBelow,
		Me.HomeCityNumber,
 ()
 ()
 ()
 ()
 ()
 52
		Me.UnitHomeCityName, Me.UnitLocation,
		Me.UnitNationNumber, Me.UnitNation, Me.UnitNear}
		Me.UnitNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitNumber.Format
 = ""
		Me.UnitNumber.FormatInfo = Nothing
		Me.UnitNumber.HeaderText = "Unit #"
		Me.UnitNumber.MappingName
 = "UnitNumber"
		Me.UnitNumber.ReadOnly = True
		Me.UnitNumber.Width
 "UnitVertCoord
 = 45
		Me.UnitVertCoord.Alignment
		Me.UnitVertCoord.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.UnitVertCoord.FormatInfo
		Me.UnitVertCoord.HeaderText
 = Nothing
 = "Vert Coord"
		Me.UnitVertCoord.MappingName = "VertCoord"
		Me.UnitVertCoord.ReadOnly = True
		Me.UnitVertCoord.Width
 t
 *UnitHorizCoord
 f
 = 70
		Me.UnitHorizCoord.Alignment
		Me.UnitHorizCoord.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.UnitHorizCoord.FormatInfo
		Me.UnitHorizCoord.HeaderText
 = Nothing
 = "Horiz Coord”
		Me.UnitHorizCoord.MappingName = "HorizCoord"
		Me.UnitHorizCoord.ReadOnly = True
		Me.UnitHorizCoord.Width
 H
 ‘Veteran
 ¥
		Me.Veteran.Alignment
		Me.Veteran.FalseValue
		Me.Veteran.HeaderText
 = 70
 = System.Windows.Forms.HorizontalAlignment.Center
 False
 "Vet"
		Me.Veteran.MappingName = "Veteran"
		Me.Veteran.NullValue
 = CType (resources.GetObject
 ("Veteran.NullValue"),
 53
 System.DBNull ¢
		Me.Veteran.ReadOnly = True
		Me.Veteran.TrueValue
 = True
		Me.Veteran.Width
 ¥
 'UnitType
 1
		Me.UnitType.Format
 = 25
 = ""
		Me.UnitType.FormatInfo
		Me.UnitType.HeaderText
		Me.UnitType.MappingName
 = Nothing
 = "Unit Type"
 = "UnitType"
		Me.UnitType.ReadOnly = True
		Me.UnitType.Width
 ¥
 *UnitOwnerColour
 f
		Me.UnitOwnerColour.Format
 = 85
 = ""
		Me.UnitOwnerColour.FormatInfo
		Me.UnitOwnerColour.HeaderText
 = Nothing
 = "Colour"
		Me.UnitOwnerColour.MappingName = "OwnerColour"
		Me.UnitOwnerColour.ReadOnly = True
		Me.UnitOwnerColour.Width
 £
 "HitPoints
 ,
		Me.HitPoints.Alignment
		Me.HitPoints.Format
		Me.HitPoints.FormatInfo
		Me.HitPoints.HeaderText
 = 60
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 I
 ¥Hlit Peints"
		Me.HitPoints.MappingName = "HitPoints"
		Me.HitPoints.ReadOnly
		Me.HitPoints.Width
 = True
 = 50
 §
 "UnitCommodity
 H
		Me.UnitCommodity.Format
 = ""
		Me.UnitCommodity.FormatInfo = Nothing
		Me.UnitCommodity.HeaderText
 = "Commodity"
		Me.UnitCommodity.MappingName
 = "UnitCommodity"
		Me.UnitCommodity.ReadOnly = True
		Me.UnitCommodity.Width
 "UnitOrders
 ¥
		Me.UnitOrders.Format
		Me.UnitOrders.FormatInfo
		Me.UnitOrders.HeaderText
 = 65
 = ""
 = Nothing
 = "Order gM
		Me.UnitOrders.MappingName = "UnitOrders"
		Me.UnitOrders.ReadOnly = True
		Me.UnitOrders.width
 "UnitvVertGotoCoords
 Hi
 = 120
		Me.UnitVertGotoCoords.Alignment
 =
 System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitVertGotoCoords.Format
		Me.UnitVertGotoCoords.FormatInfo
		Me.UnitVertGotoCoords.HeaderText
		Me.UnitVertGotoCoords.MappingName
		Me.UnitVertGotoCoords.ReadOnly
		Me.UnitVertGotoCoords
 t
 '"UnitHorizGotoCoords
 {
		Me.UnitHorizGotoCoords.Alignment
		Me.UnitHorizGotoCoords.Format
		Me.UnitHorizGotoCoords.FormatInfo
		Me.UnitHorizGotoCoords.HeaderText
 = ""
 =
 Width = 60
 = "
 = Nothing
 = "Goto Vert”
 = "VertGotoCoords"
 True
 = System.Windows.Forms.HorizontalAlignment.Right
 TY
 Nothing
 = "Goto Horiz"
		Me.UnitHorizGotoCoords.MappingName = "HorizGotoCoord"
		Me.UnitHorizGotoCoords.ReadOnly
		Me.UnitHorizGotoCoords.Width
 |;
 "HomeCityNumber
 f
		Me.HomeCityNumber.Alignment
		Me.HomeCityNumber.Format
 =
 = 60
 = Sys
 = ""
 True
 tem.Windows.Forms.HorizontalAlignment.Right
		Me.HomeCityNumber.FormatInfo = Nothing
		Me.HomeCityNumber.HeaderText
 = "Home City #"
		Me.HomeCityNumber .MappingName = "HomeCityNumber"
		Me.HomeCityNumber.ReadOnly
 = True
		Me.HomeCityNumber.Width
 H
 'UnitAbove
 ¥
 = 75
		Me.UnitAbove.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitAbove.Format
 = ""
		Me.UnitAbove.FormatInfo
		Me.UnitAbove.HeaderText
		Me.UnitAbove.MappingName
 = Nothing
 "Unit # Above"
 = "UnitAbove"
		Me.UnitAbove.ReadOnly = True
		Me.UnitAbove.Width
 f
 "UnitBelow
 §
 = 75
		Me.UnitBelow.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitBelow.Format
 = ""
		Me.UnitBelow.FormatInfo = Nothing
		Me.UnitBelow.HeaderText
		Me.UnitBelow.MappingName
 = "Unit # Below"
 = "UnitBelow"
		Me.UnitBelow.ReadOnly = True
		Me.UnitBelow.Width
 §
 "UnitHomeCityName
 ¥
 = 75
		Me.UnitHomeCityName.Format
		Me.UnitHomeCityName.FormatInfo
		Me.UnitHomeCityName.HeaderText
 = ""
 =
 =
 Nothing
 54
 "Home City Name"
		Me.UnitHomeCityName.MappingName
		Me.UnitHomeCityName.Width
 = "UnitHomeCityName"
		Me.UnitHomeCityName.ReadOnly
 = True
 = 101
 i
 "UnitLocation
 i
		Me.UnitLocation.Format
 = ""
		Me.UnitLocation.FormatInfo
 = Nothing
		Me.UnitLocation.HeaderText = "Location"
		Me.UnitLocation.MappingName = "UnitLocation"
		Me.UnitLocation.ReadOnly
		Me.UnitLocation.Width
 ti
 "UnitNationNumber
 = True
 = 55
 59
		Me.UnitNationNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitNationNumber.Format
 = ""
		Me.UnitNationNumber.FormatInfo
 = Nothing
		Me.UnitNationNumber.HeaderText = "Nation #"
		Me.UnitNationNumber.MappingName = "UnitNationNumber”
		Me.UnitNationNumber.ReadOnly = True
		Me.UnitNationNumber.Width
 ¥
 "UnitNation
 ¥
		Me.UnitNation.Format
		Me.UnitNation.FormatInfo
		Me.UnitNation.HeaderText
 = 50
 = ""
 = Nothing
 = "Nationality"
		Me.UnitNation.MappingName = "UnitNation"
		Me.UnitNation.ReadOnly
		Me.UnitNation.Width
 = True
 = 75
 t
 'UnitNear
 ¥
		Me.UnitNear.Format
 = ""
		Me.UnitNear.FormatInfo
		Me.UnitNear.HeaderText
 = Nothing
 = "Near"
		Me.UnitNear.MappingName = "UnitNear"
		Me.UnitNear.ReadOnly = True
		Me.UnitNear.Width
 End Sub
 #End Region
 #Region "Wonders"
 Friend
 Friend
 Friend
 Friend
 Friend
 Friend
 WithEvents
 WithEvents
 WithEvents
 WithEvents
 = 202
 WonderNumber As CIV _II Extractor.DataGridColouredTextBoxColumn
 WonderName As CIV II Extractor.DataGridColouredTextBoxColumn
 WonderEra As CIV _
 TI _Extractor.
 WonderCityName
 DataGridColouredTextBoxColumn
 As CIV. ITI Extractor.DataGridColouredTextBoxColumn
 WithEvents
 WithEvents
 Friend
 Private
 WithEvents
 WonderBuilt
 WonderDestroyed
 As CIV _II Extractor.
 DataGridColouredBoolColumn
 As CIV _II Extractor.DataGridColouredboolColumn
 WonderCityColour
 Ag CIV II Extractor.DataGridColouredTextBoxColumn
 Sub SetupWonderColumns (ByRef resources
		Me.WonderNumber
 As System.Resources.ResourceManager)
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.WonderName
		Me
 .WonderEra
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.WonderCityName
		Me.WonderBuilt
		Me.WonderDestroyed
		Me.WonderCityColour
 mWondersCS
 —
 —
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV_II Extractor.DataGridColouredBoolColumn
 = New CIV. II Extractor.DataGridColouredBoolColumn
 = New CIV IT_Extractor.DataGridColouredTextBoxColumn
 New System.Windows.Forms.DataGridColumnStyle()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 {Me.WonderNumber,
		Me.WonderName, Me.WonderEra, Me.WonderCityName, Me.WonderBuilt,
		Me.WonderDeatroyed, Me.WonderCityColour)
 ‘Wo
 f
 nderNumber
		Me.WonderNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.WonderNumpber.Format
 = ""
		Me.WonderNumber.FormatInfo = Nothing
		Me.WonderNumber.HeaderText
 = "Wonder #"
		Me. WonderNumber .MappingName = "WonderNumber"
		Me.WonderNumber.ReadOnly = True
		Me.WonderNumber.Width
 f
 'WonderName
 ¥
		Me.WonderName.Format
 = 60
 = ""
		Me.WonderName.FormatInfo = Nothing
		Me.WonderName.HeaderText
		Me.WonderName.MappingName
 = "Wonder Name"
 = "WonderName"
		Me.WonderName.ReadOnly = True
		Me.WonderName.Width = 135
 H
 'WonderEra
 §
		Me.WonderEra.Format
 = ""
		Me.WonderEra.FormatInfo = Nothing
		Me.WonderEra.HeaderText = "Era"
		Me.WonderEra.MappingName = "WonderEra”
		Me.WonderEra.ReadOnly = True
		Me.WonderEra.Width
 f
 "WonderCityName
 ¥
		Me.WonderCityName.Format
 = 75
 = ""
		Me.WonderCityName.FormatInfo = Nothing
		Me.WonderCityName.HeaderText = "Located In"
		Me.WonderCityName.MappingName = "WonderCityName"
		Me.WonderCityName.ReadOnly
		Me.WonderCityName.Width
 = True
 = 125
 f
 'WonderBuilt
		Me.WonderBuilt.Alignment
		Me.WonderBuilt.FalseValue = False
 II
 Extractor\Forml.vb
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.WonderBuilt.HeaderText = "Built"
		Me.WonderBuilt.
		Me.WonderBuilt.NullValue
		Mem ingens
 System.DBNull)
		Me.WonderBuilt.ReadOnly
		Me.WonderBuilt.TrueValue
		Me.WonderBuilt.Width
 "WonderDestroyed
 §
		Me.WonderDestroyed.Alignment
		Me.WonderDestroyed.FalseValue
 = 60
 = "WonderBuilt"
 = CType (resources.
 =
 True
 — ——
 True
 co
 GetObject
 ("WonderBuilt.
 Nullvalue"),
 System.Windows.Forms.HorizontalAlignment.Center
 =
 )
		Me.WonderDestroyed.HeaderText
		Me.WonderDestroyed.MappingName
		Me.WonderDestroyed.NullValue
 System.DBNull)
		Me.WonderDestroyed.ReadOnly
		Me.WonderDestroyed.TrueValue
		Me.WonderDestroyed.Width
 ¥
 'WonderCityColour
 f
		Me.WonderCityColour.Format
		Me.WonderCityColour.FormatInfo
		Me.WonderCityColour.HeaderText
		Me.WonderCityColour.MappingName
		Me.WonderCityColour.Width
 End Sub
 #End Region
 #Region "Triumphs"
 Friend WithEvents
 Friend WithEvents
 TriumphNumber
 =
 60
 = 75
 As
 TriumphNationNumb
 Friend WithEvents
 Friend WithEvents
 TriumphNation
 TriumphYear
 As
 T
 TY
 False
 —
 r=
 "Destroyed"
 "WonderDestroyed"”
 CType (resources.GetObject
 rue
 True
 —
 —
 Nothing
 = "Colour"
 = "WonderCityColour"
 CIV_IT Extractor.
 ("WonderDestroyed.NullvValue"
 DataGridColouredTextBoxColumn
 er As CIV_ ITExtractor.DataGridColouredTextBoxColumn
 CIV II Extractor.
 DataGridColouredTextBoxColumn
 As CIV TIT Extractor.DataGridColouredTextBoxColumn
 56
 v4
 ¢
 Friend
 Friend
 WithEvents
 WithEvents
 TriumphTurn
 TriumphNationColour
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 Private
 Sub SetupTriumphColumns (ByRef resources As System.Resources.ResourceManager)
		Me.TriumphNumber
		Me.TriumphNationNumber
		Me.TriumphNation
		Me.TriumphYear
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me. TriumphTurn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.TriumphNationColour
 mTriumphsCS
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New System.Windows.Forms.DataGridColumnStyle()
 ()
 ()
 ()
 ()
 ()
 ()
 {Me.TriumphNumber,
		Me. TriumphNationNumber, Me.TriumphNation, Me.TriumphYear, Me.TriumphTurn,
		Me.TriumphNationColour}
 J
 'TriumphNumber
 ¥
		Me.TriumphNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.TriumphNumber.Format
 = ""
		Me.TriumphNumber.FormatInfo = Nothing
		Me. TriumphNumber.HeaderText = "Triumph#"
		Me. TriumphNumber.MappingName = "TriumphNumber"
		Me.TriumphNumber.ReadOnly = True
		Me. TriumphNumber. .Width = 60
 H
 "TriumphNationNumber
 f
		Me.TriumphNationNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.TriumphNationNumber.Format
 = ""
		Me.TriumphNationNumber.FormatInfo
		Me. TriumphNationNumber.HeaderText
		Me. TriumphNationNumber.MappingName
		Me. TriumphNationNumber.ReadOnly
		Me.TriumphNationNumber.Width
 §
 "TriumphNation
 §
		Me.TriumphNation.Format
 = ""
		Me.TriumphNation.FormatInfo
 = 60
 Nothing
 "Nation#"
 = "TriumphNationNumber"
 = True
 = Nothing
		Me.TriumphNation.HeaderText = "Nation Triumphed”
		Me.TriumphNation.MappingName = "TriumphNation"
		Me.TriumphNation.ReadOnly = True
		Me. TriumphNation.Width
 ¥
 '"TriumphYear
 = 100
		Me.TriumphYear.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.TriumphYear.Format
 = ""
		Me.TriumphYear.FormatInfo = Nothing
		Me.TriumphYear.HeaderText = "Year"
		Me.TriumphYear.MappingName
 = "TriumphYear"
		Me.TriumphYear.ReadOnly = True
		Me.TriumphYear.Width
 '"TriumphTurn
 = 60
		Me.TriumphTurn.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.TriumphTurn.Format
 = ""
		Me.TriumphTurn.FormatInfo
 = Nothing
		Me.TriumphTurn.HeaderText = "Turn#"
		Me.TriumphTurn.MappingName
 = "TriumphTurn"
		Me.TriumphTurn.ReadOnly = True
		Me.TriumphTurn.wWidth
 'TriumphNationColour
		Me.TriumphNationColour.Format
 = 60
 = ""
		Me.TriumphNationColour.FormatInfo
		Me.TriumphNationColour.HeaderText
 = Nothing
 = "Colour™
		Me.TriumphNationColour.MappingName = "TriumphNationColour"
		Me.TriumphNationColour.ReadOnly = True
		Me.TriumphNationColour.
 Width = 75
 57
 End Sub
 #End Region
 #Region "Treaties"
 Friend WithEvents
 FromCivColourNumber As CIV II Extractor.DataGridColouredTextBoxColumn
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Private
 ToCivColour As CIV IT_Extractor. DataGridColouredTextBoxColumn
 ToCivNation
 Contact
 CeaseFire
 As CTW: ;IT _Extractor.
 As CIV II Extractor.
 DataGridColouredTextBoxColumn
 DataGridColouredBoolColumn
 As CIV II Extractor.DataGridColouredBoolColumn
 Peace As CIV II Extractor.DataGridColouredBoolColumn
 Alliance
 Vendetta
 As CIV II Extractor.DataGridColouredBoolColumn
 As CIV II Extractor.DataGridColouredBoolColumn
 Embassy As CIV II Extractor.DataGridColouredBoolColumn
 War As CIV_II Extractor.DataGridColouredBoolColumn
 Attitude
 As CIV II Extractor.DataGridColouredTextRoxColumn
 LastContactTurn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 LastContactYear
 Sub SetupTreatiesColumns
		Me.FromCivColourNumber
 As CIV II Extractor.DataGridColouredTextBoxColumn
 (ByRef resources
 As System.Resources.ResourceManager)
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.ToCivColour
		Me.ToCivNation
		Me.Contact
		Me.CeaseFire
		Me.Peace
 =
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV ITI Extractor.DataGridColouredBoolColumn
 = New CIV TIT _Extractor.
		Me.Alliance
 DataGridColouredBoolColumn
 = New eTV ITExtractor.DataGridColouredBoolColumn
		Me.Vendetta
		Me.Embassy
		Me.War
 = New CIV II Extractor.
 DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
 = New CIV II Extractor.DataGridColouredBoolColumn
		Me.Attitude
 = New CIV ITIExtractor.DataGridColouredTextBoxColumn
		Me.LastContactTurn
 ()
 {)
 ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.LastContactYear
 m TreatiesCS
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 = New System.Windows.Forms.DataGridColumnStyle()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 {Me.ToCivColour,
		Me. ToCivColour, Me .ToCivNation,
		Me.Alliance,
		Me.Vendetta,
		Me.Contact,
		Me.CeaseFire,
		Me.Embassy, Me.War, Me.Attitude,
		Me.LastContactYear}
 EH
 'FromCivColourNumber
 ¢
		Me.Peace,
		Me.LastContactTurn,
 ()
		Me.FromCivColourNumber.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.FromCivColourNumber.Format
 = ""
		Me.FromCivColourNumber.FormatInfo = Nothing
		Me.FromCivColourNumber.HeaderText
		Me.FromCivColourNumber.MappingName
 = "From Col#"
 = "FromCivColourNumber"
		Me.FromCivColourNumber.ReadOnly = True
		Me.FromCivColourNumber.Width
 LH
 I TalivColour
 M
 Ma .ToCivColour.Format
		Me.ToCivColour.FormatInfo
		Me.ToCivColour.HeaderText
		Me.ToCivColour.MappingName
 = 60
 = "M
 = Nothing
 = "To Colour”
 = "ToCivColour"
		Me.ToCivColour.ReadOnly = True
		Me.ToCivColour.Width
 tH
 "ToCivNation
 §
		Me.ToCivNation.Format
 = 75
 = ""
		Me.ToCivNation.FormatInfo
		Me.ToCivNation.HeaderText
		Me.ToCivNation.MappingName
 = Nothing
 = "To Nation”
		Me.ToCivNation.ReadOnly = True
 = "ToCivNation"
		Me.ToCivNation.Width
 H
 "Contact
 |
		Me.Contact.Alignment
 = 60
 = System.Windows.Forms.HorizontalAlignment.Center
 58
		Me.Contact.FalseValue = False
		Me.Contact.HeaderText
 = "Con"
		Me.Contact.MappingName = "Contact"
		Me.Contact.NullvValue
 = CType (resources.GetObject
 ("Contact.NullValue"),
		Me.Contact.ReadOnly = True
		Me.Contact.TrueValue
 = True
		Me.Contact.Width
 '"CeaseFire
 H
 = 45
		Me.CeaseFire.Alignment
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.CeaseFire.FalseValue = False
		Me.CeaseFire.HeaderText
 = "C FEF"
		Me.CeaseFire.MappingName = "CeaseFire"
		Me.CeaseFire.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.CeaseFire.ReadOnly = True
		Me.CeaseFire.TrueValue
 = True
		Me.CeaseFire.Width
 ¥
 ‘Peace
 §
 = 45
 ("CeaseFire.Nullvalue"),
		Me.Peace.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me. Peace.FalseValue
 = False
		Me.Pecace.HeaderText = "Pce”
		Me.Peace.MappingName = "Peace"
		Me.Peace.NullValue
 = CType (resources.GetObject
		Me.Peace.ReadOnly = True
		Me.Peace.TrueValue = True
		Me.Peace.Width
 14
 ‘Alliance
 = 45
		Me.Alliance.Alignment
 ("Peace.NullValue"),
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Alliance.FalseValue = False
		Me.Alliance.HeaderText
 "All"
		Me.Alliance.MappingName = "Alliance"
		Me.Alliance.NullValue
 DBNull)
 = CType (resources.GetObject("Alliance.Nullvalue"),
		Me.Alliance.ReadOnly = True
		Me.Alliance.TrueValue
		Me.Alliance.Width
 ¥
 ‘Vendetta
 ¥
 = 45
		Me.Vendetta.Alignment
 = True
 = System.Windows.Forms.HorizontalAlignment.Center
		Me.Vendetta.FalseValue = False
		Me.Vendetta.HeaderText
 = "Ven"
		Me.Vendetta.MappingName = "Vendetta"
		Me.Vendetta.NullValue
 DBNull)
 = CType (resources.GetObject
		Me.Vendetta.ReadOnly = True
		Me.Vendetta.TrueValue
		Me.Vendetta.Width
 t
 ‘Embassy
 §
 = 45
 = True
 ("Vendetta.NullValue"),
		Me.Embassy.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.Embassy.FalseValue = False
		Me.Embassy.HeaderText
 = "Emb"
		Me.Embassy.MappingName = "Embassy"
		Me.Embassy.NullValue
 = CType (resources.GetObject
 ("Embassy.NullValue"),
 System.DBNull«
 System.
 System.DBNull)
 System.
 System.
 59
 System.DBNullw«
		Me.Embassy.ReadOnly = True
		Me. Embassy.TrueValue = True
		Me. Embassy.Width
 ¥
 "War
 1
 = 45
		Me.War.Alignment = System.Windows.Forms.HorizontalAlignment.Center
		Me.War.FalseValue = False
		Me.War.HeaderText
 "War"
		Me.War.MappingName = "War"
 «
 v4
 ve
		Me.War.NullValue
 = CType (resources.GetObject
		Me.War.ReadOnly = True
		Me.War.wWidth
		Me.War.TrueValue = True
 = 43
 ¥
 "Attitude
 i
		Me.Attitude.Alignment
 ("War.NullValue"),
 System.DBNull)
 = System.Windows.Forms.HorizontalAlignment.Right
		Me.Attitude.Format
		Me.Attitude.FormatInfo
 = ""
 = Nothing
		Me.Attitude.HeaderText = "Attitude"
		Me.Attitude.MappingName = "Attitude"
		Me.Attitude.ReadOnly
		Me.Attitude.Width
 ‘LastContactT
 f
 urn
 = True
 = 75
		Me.LastContactTurn.Alignment
		Me.LastContactTurn.Format
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
		Me.LastContactTurn.FormatInfo
		Me.LastContactTurn.HeaderText
 = Nothing
 = "Last Cont Turn#"
		Me.LastContactTurn.MappingName = "LastContactTurn"
		Me.LastContactTurn.ReadOnly
		Me.LastContactTurn.Width
 t
 'TLastContactYear
 §
		Me.LastContactYear.Alignment
		Me.LastContactYear.Format
		Me.LastContactYear.FormatInfo
		Me.LastContactYear.HeaderText
 = True
 = 100
 = System.Windows.Forms.HorizontalAlignment.Right
 = ""
 = Nothing
 = "Last Contact Year"
		Me.LastContactYear.MappingName = "LastContactYear"
		Me.LastContactYear.ReadOnly
		Me.LastContactYear.Width
 End Sub
 #End Region
 #Region "UnitNationTotals"
 Friend WithEvents
 = True
 = 75
 UnitType2 As CIV ITExtractor.DataGridColouredTextBoxColumn
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Friend WithEvents
 Private
 NationNumber2
 Nation2
 As CIV II Extractor.DataGridColouredTextBoxColumn
 As CIV II Extractor.DataGridColouredTextBoxColumn
 UnitNationColour
 As CIV II Extractor.DataGridColouredTextBoxColumn
 UnitNationCount
 As CIV II Extractor.DataGridColouredTextBoxColumn
 Sub SetupUnitNationTotalsColumns
 ResourceManager)
		Me.UnitType2
 (ByRef resources
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.NationNumber2
 As System.Resources.
 ()
 = New CIV II Extractor.DataGridColouredTextBoxColumn
		Me.Nation2
 = New CIV ITI Extractor.DataGridColouredTextBoxColumn
		Me.UnitNationColour
 = New CIV II Extractor.DataGridColouredTextBoxColumn
 ()
		Me.UniltNationCount = New CIV II Extractor.DataGridColouredTextBoxColumn
 t
 '"UnitType?2
		Me.UnitType?2.
		Me.UnitType?.
		Me.UnitTypeZ.
		Me.UnitType2.
		Me.UnitType2.
		Me.UnitType?2.
 H
 '‘NationNumberZ
 Format
 = ""
 FormatInfo = Nothing
 HeaderText = "Unit Type"
 MappingName = "UnitType"
 ReadOnly = True
 width
 = 85
 ()
		Me.NationNumber2.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.NationNumber2.Format
 = ""
		Me.NationNumber2.FormatInfo = Nothing
		Me.NationNumber?2 .HeaderText = "Nation#"
		Me.NationNumber?2 .MappingName = "NationNumber"
		Me.NationNumberz.ReadOnly = True
		Me.NationNumber2.Width
 ()
 ()
 60
 = 50
 "Nation?
 £
		Me.Nation2.Format
 = ""
		Me.Nation2.FormatInfo
 = Nothing
		Me.Nation2.HeaderText
		Me.Nation?2.MappingName = "Nation"
 = "Nation"
		Me.Nation2.ReadOnly = True
		Me.Nation2.Width
 f
 "UnitNationColour
		Me.UnitNationColour.Format
 = 75
		Me.UnitNationColour.FormatInfo
 = ""
 = Nothing
		Me.UnitNationColour .HeaderText = "Colour"
		Me.UnitNationColour.MappingName = "UnitNationColour"
		Me.UnitNationColour.ReadOnly
		Me.UnitNationColour.Width
 ¥
 "UnitNationCount
		Me.UnitNationCount.Alignment
 = True
 = 60
 = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitNationCount.Format
 = ""
		Me.UnitNationCount.FormatInfo
 = Nothing
		Me.UnitNationCount.HeaderText = "# Of Units"
		Me.UnitNationCount.MappingName = "UnitNationCount"
		Me.UnitNationCount.ReadOnly
		Me.UnitNationCount.Width
 = True
 = 65
 m UnitNationTotalsCS
 {Me.UnitType2,
		Me.UnitNationCount}
 End Sub
 #End Region
 #Region "UnitType"
 = New System.Windows.Forms.DataGridColumnStyle()
		Me.NationNumber?2, Me.Nation?,
		Me.UnitNationColour,
 Friend WithEvents UnitType3d As DataGridTextBoxColumn
 Friend WithEvents UnitTypeCount As DataGridTextBoxColumn
 Private
 Sub SetupUnitTypeColumns
 (ByRef resources
		Me.UnitType3 = New DataGridTextBoxColumn ()
		Me.UnitTypeCount
 mUnitTypeCS
 = New DataGridTextBoxColumn
 As System.Resources.ResourceManager)
 ()
 = New System.Windows.Forms.DataGridColumnStyle()
		Me.UnitTypeCount}
 t
 "UnitTypel
 H
		Me.UnitType3.Format
 = ""
		Me.UnitType3d.FormatInfo = Nothing
		Me.UnitType3.HeaderText
 = "Unit Type"
		Me.UnitType3.MappingName = "UnitType"
		Me.UnitType3.ReadOnly = True
		Me.UnitType3.Width
 'UnitTypelount
 = 895
 {Me.UnitType3,
		Me.UnitTypeCount.Alignment = System.Windows.Forms.HorizontalAlignment.Right
		Me.UnitTypeCount.Format
 = ""
		Me.UnitTypeCount.FormatInfo
		Me.UnitTypeCount.HeaderText
 = Nothing
 = "# Of Units"
		Me.UnitTypeCount.MappingName = "UnitTypeCount"
		Me.UnitTypeCount.ReadOnly = True
		Me.UnitTypeCount.wWidth
 End Sub
 #End Region
 #End Region
 Private
 = 65
 Sub SetupDBR()
 '
 Create DataRelations
 '
 CITIES
 to
 and add them to the DataSet.
 UNITS on CITYNUMBER
 m drl
 = New DataRelation("UnitsOwned",
 dsCiv.Units.HomeCityNumberColumn)
 dsCiv.Relations.Add (m drl)
 '
 CIVILIZATION
 m dr2
 TO TREATY on CIV COLOUR
 = New DataRelation("Treaty",
 dsCiv.Cities.CityNumberColumn,
 dsCiv.Civilization.CivColourNumberColumn,
 dsCiv.Treaty.FromCivColourNumberColumn)
 dsCiv.Relations.Add(m
 dr2)
 " NATION TO CIV on COLOUR
 m dr3
 = New DataRelation("CivColour",
 dsCiv.Civilization.CivColourNumberColumn)
 dsCiv.Relations.Add(m dr3)
 '" NATION TO TRIUMPH on NATION NUMBER
 m dr4
 = New DataRelation
 ("NationTriumphed",
 dsCiv.Triumph.TriumphNationNumberColumn)
 dsCiv.Relations.Add(m dr4)
 "UNIT
 dsCiv.Nation.NationColourNumberColumn,
 dsCiv.Nation.NationNumberColumn,
 TYPE to UNITNATIONTOTALS on UNITTYPE
 m dr5
 = New DataRelation("Nations",
 dsCiv.UnitNationTotals.UnitTypeColumn)
 dsCiv.Relations.Add(m dr5)
 dsCiv.UnitType.UnitTypeColumn,
 " NATION TO UNITNATIONTOTALS on NATIONNUMBER
 mdr6
 = New DataRelation
 ("UnitCounts"”,
 dsCiv.UnitNationTotals.NationNumberColumn)
 dsCiv.Relations.Add(m dr6)
 '
 UNIT TYPE to UNITS on UNITTYPE
 mdr7
 = New DataRelation("UnitsForType",
 dsCiv.Units.UnitTypeColumn)
 dsCiv.Relations.Add (m dr7)
 dsCiv.Nation.NationNumberColumn,
 dsCiv.UnitType.UnitTypeColumn,
 " UNITNATIONTOTALS to UNITS on UNITTYPE & NATIONNUMBER combined
 Dim PKey8(1l), CKey8(l) As DataColumn
 PKey8 (1) = dsCiv.UnitNationTotals.UnitTypeColumn
 PKey8 (0) = dsCiv.UnitNationTotals.NationNumberColumn
 CKey8(1l) = dsCiv.Units.UnitTypeColumn
 CKey8 (0) = dsCiv.Units.UnitNationNumberColumn
 mdr8
 = New DataRelation("UnitsForTypeAndNation",
 dsCiv.Relations.Add (m dr8)
 " MAPCELL to CITIES
 Dim PKey9(1l),
 CKey9(l)
 As DataColumn
 PKey9 (0) = dsCiv.MapCell.MapHorizCoordColumn
 PKey9 (1) = dsCiv.MapCell.MapVertCoordColumn
 CKey9 (0) = dsCiv.Cities.HorizCoordColumn
 CKey9 (1) = dsCiv.Cities.VertCoordColumn
 m dr9 = New DataRelation("City
 dsCiv.Relations.Add (m dr9)
 '
 MAPCELL to UNITS
 at
 Cell",
 |
 Dim PKeyl0{(1l), CKeyl0O(l) As DataColumn
 PKey9,
 PKeyl0(0) = dsCiv.MapCell.MapHorizCoordColumn
 PKey10 (1)
 CKeyl0 (0)
 il
 li
 dsCiv.MapCell.MapVertCoordColumn
 dsCiv.Units.HorizCoordColumn
 CKeyl0(1l) = dsCiv.Units.VertCoordColumn
 m drl0
 = New DataRelation
 ("Units
 dsCiv.Relations.Add (m drl0)
 End Sub
 #Region "Table Styles”
 Private
 Sub SetupDGTS ()
 SetupNationsDGTS
 SetupCivsDGTS
 SetupCitiesDGTS
 SetupUnitsDGTS
 SetupUnitCountsDGTS
 SetupWondexrsDGTS
 SetupTriumphsDGTS
 End Sub
 #Region "Nations"
 ()
 ()
 ()
 ()
 ()
 ()
 ()
 at
 Cell"™,
 PKey8,
 CKey9)
 PKeylO,
 CKeylO)
 CKey8)
 Friend WithEvents dgNatTSNat As System.Windows.Forms.DataGridTableStyle
 Friend WithEvents dgNatTS3Civs As System.Windows.Forms.DataGridTableStyle
 Friend WithEvents dgNatTSTreaties As System.Windows.Forms.DataGridTableStyle
 62
 Friend WithEvents dgNatTSTriumphs As System.Windows.Forms.DataGridTablestyle
Flowers\My...
  Friend WithEvents dgNatTSUnitNationTotals As System.Windows.Forms.DataGridTableStyle
 Friend WithEvents dgNatTSUnits As System.Windows.Forms.DataGridTableStyle
 Private
 '
 Sub SetupNationsDGTS ()
 dgNations
 dgNatTsSNat (Nation)
		Me.dgNatTSNat
		Me.dgNat.TableStyles.AddRange
 = New System.Windows.Forms.DataGridTableStyle
 {Me .dgNatTSNat})
		Me.dgNatTSNat.DataGrid
 = Me.dgNat
		Me.dgNatTSNat.GridColumnStyles.AddRange (m NationsCS)
 ()
 (New System.Windows.Forms.DataGridTableStyle()
		Me.dgNatTSNat.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgNatTSNat.MappingName
 = "Nation"
		Me.dgNatTSNat.ReadOnly = True
 :
		Me.dgNatTSCivs
 dgNatT3Civs Relation CivColor [m dr3] Nation to (Civilization)
 = New System.Windows.Forms.DataGridTableStyle()
		Me.dgNat.TableStyles.AddRange
 (New System.Windows.Forms.DataGridTableStyle()
 dgNatTsSCivs})
		Me.dgNatTSCivs.DataGrid
 = Me.dgNat
		Me.dgNatTSCivs.GridColumnStyles.AddRange
		Me.dgNatTSCivs.HeaderForeColor
 |
 (m CivsCS)
 = System.Drawing.SystemColors.ControlText
		Me.dgNatTSCivs.MappingName
 = "Civilization"
		Me.dgNatTSCivs.ReadOnly = True
 :
 dgNatTSTreaties Relation Treaty [mdr2] Civilization
		Me.dgNatTSTreaties
 dgNat.TableStyles.AddRange
 = New System.Windows.Forms.DataGridTableStyle
 to (Treaty)
 (New System.Windows.Forms.DataGridTableStyle()
 dgNatTSTreaties})
		Me.dgNatTSTreaties.DataGrid
 = dgNat
		Me.dgNatTSTreaties.GridColumnStyles.AddRange
		Me.dgNatTSTreaties.HeaderForeColor
 =
		Me.dgNatTSTreaties.MappingName = "Treaty"
 (m TreatiesCS)
 ()
 System.Drawing.SystemColors.ControlText
		Me.dgNatTSTreaties.ReadOnly
 = True
		Me.dgNatTSTreaties.RowHeadersVisible
 ,
 dgNatTSTriumphs
 Relation
 = True
 NationTriumphed
		Me.dgNatTSTriumphs
 [m dr4]
 = New System.Windows.Forms.DataGridTableStyle
		Me.dgNat.TableStyles.AddRange
 Nation to (Triumph)
 ()
 (New System.Windows.Forms.DataGridTablesStyle
 dgNatTSTriumphs})
		Me.dgNatTSTriumphs.DataGrid
 = Me.dgNat
		Me.dgNatTSTriumphs.GridColumnStyles.AddRange
		Me.dgNatTSTriumphs.HeaderForeColor
		Me.dgNatTSTriumphs.MappingName
		Me.dgNatTSTriumphs.ReadOnly = True
		Me.dgNatTSTriumphs.RowHeadersVisible = False
 dgNatTSUnitNationTotals
 (m TriumphsCS)
 =
 System.Drawing.SystemColors.ControlText
 = "Triumph"
 {Me.
 {Me.
 ()
 {Me.
 63
 Relation UnitCounts [mdré] Nation to (UnitNationTotalswe
		Me.dgNatTSUnitNationTotals
		Me.dgNat.TableStyles.AddRange
 {dgNatTsSUnitNationTotals})
 = New System.Windows.Forms.DataGridTableStyle()
 (New System.Windows.Forms.DataGridTableStyle
		Me.dgNatTSUnitNationTotals.DataGrid
 = Me.dgNat
		Me.dgNatTSUnitNationTotals.GridColumnStyles.AddRange
		Me.dgNatTSUnitNationTotals.HeaderForeColor
 =
 (m UnitNationTotalsCS)
 System.Drawing.SystemColors.ControlText
		Me.dgNatTSUnitNationTotals.MappingName = "UnitNationTotals"
		Me.dgNatTSUnitNationTotals.ReadOnly
		Me.dgNatTSUnitNationTotals.RowHeadersVisible
 ¥
 = True
 =
 True
 ()
 dgNatTST3Units Relation UnitsForTypeAndNation [mdr8] UnitNationTotals to (Units
		Me.dgNatTSUnits = New System.Windows.Forms.DataGridTableStyle
		Me.dgNat.TableStyles.AddRange
 ()
 (New System.Windows.Forms.DataGridTableStyle()
 dgNatTSUnits})
		Me.dgNatTSUnits.DataGrid
 = Me.dgNat
		Me.dgNatTSUnits.GridColumnStyles.AddRange
 (mUnitsCS)
		Me.dgNatTsSUnits.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgNatTSUnits.MappingName = "Units"
		Me.dgNatTSUnits.RowHeadersVisible
		Me.dgNatTSUnits.ReadOnly = True
 End Sub
 #End Region
 = False
 #Region "Civs"
 Friend WithEvents dgCivsTSCivs As System.Windows.Forms.DataGridTableStyle
 Friend WithEvents dgCivsTSTreaties As System.Windows.Forms.DataGridTableStyle
 {Me.
 v
 Private
 Sub SetupCivsDGTS ()
 " dgCivs
 ¥
 dgCivsTS8Civs (Civilization)
		Me
		Me
 .dgCivsTSCivs
 = New System.Windows.Forms.DataGridTableStyle
 .dgCivs.TableStyles.AddRange
 ()
 (New System.Windows.Forms.DataGridTableStyle()
 dgCivsTSCivs.DataGrid = Me.dgCivs
 dgCivsTSCivs})
		Me.
		Me.
		Me
		Me
 dgCivsTSCivs.GridColumnStyles.AddRange
 .dgCivsTSCivs.HeaderForeColor
 (m CivsCS)
 = System.Drawing.SystemColors.ControlText
 .dgCivsTsSCivs.MappingName
		Me
 f
 = "Civilization"
 .dgCivsTSCivs.ReadOnly = True
 dgCivsTSTreaties Relation Treaty [mdr2] Civilization
		Me
 .dgCivsTSTreaties
 dgCivs.TableStyles.AddRange
 = New System.Windows.Forms.DataGridTableStyle
 to (Treaty)
 (New System.Windows.Forms.DataGridTableStyle()
 dgCivsTSTreaties})
		Me
		Me
		Me
		Me
		Me
		Me
 .dgCivsTSTreaties.DataGrid
 = dgCivs
 .dgCivsTSTreaties.GridColumnStyles.AddRange
 (m TreatiesCS)
 .dgCivsTSTreaties.HeaderForeColor = System.Drawing.SystemColors.ControlText
 .dgCivsTSTreaties.MappingName = "Treaty"
 .dgCivsTSTreaties.ReadOnly
 = True
 .dgCivsTSTreaties.RowHeadersVisible
 End Sub
 #End Region
 #Region "Cities"
 Friend
 Friend
 WithEvents dgCitiesTSCities
 = True
 As System.Windows.Forms.DataGridTableStyle
 WithEvents dgCitiesTSUnits As System.Windows.Forms.DataGridTableStyle
 Private
 '
 '
 Sub SetupCitiesDGTS ()
 dgCities
 dgCitiesTsCities (Cities)
		Me.dgCitiesTSCities
 = New System.Windows.Forms.DataGridTableStyle()
		Me.dgCities.TableStyles.AddRange
 {Me.dgCitiesTSCities})
 (New System.Windows.Forms.DataGridTableStyle()
		Me.dgCitiesTSCities.DataGrid = Me.dgCities
		Me.dgCitiesTSCities.GridColumnStyles.AddRange (m CitiesCS)
		Me.dgCitiesTSCities.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgCitiesTSCities.MappingName
		Me.dgCitiesTSCities.ReadOnly
 = "Cities"
 = True
 dgCitiesT8Units
		Me.dgCitiesTSUnits
 Relation
 UnitsOwned [m drl] Cities to (Units)
 = New System.Windows.Forms.DataGridTableStyle
		Me.dgCities.TableStyles.AddRange
 dgCitiesTSUnits})
 (New System.Windows.Forms.DataGridTablesStyle()
		Me.dgCitiesTS8SUnits.DataGrid = Me.dgCities
		Me.dgCitiesTSUnits.GridColumnStyles.AddRange
 (m UnitsCS)
		Me.dgCitiesTSUnits.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgCitiesTSUnits.MappingName = "Units"
		Me.dgCitiesTSUnits.RowHeadersVisible = False
		Me.dgCitiesTSUnits.ReadOnly
 End Sub
 #End Region
 #Region "Units"
 Friend
 = True
 WithEvents dgUnitsTSUnits As System.Windows.Forms.DataGridTableStyle
 Private
 '
 :
 Sub SetupUnitsDGTS ()
 dgUnits
 dgUnitsT8Units (Units)
		Me.dgUnitsTSUnits
 = New System.Windows.Forms.DataGridTableStyle()
		Me.dgUnits.TableStyles.AddRange
 dgUnitsTSUnits})
		Me.dgUnitsTSUnits.DataGrid
 (New System.Windows.Forms.DataGridTableStyle()
 = Me.dgUnits
		Me.dgUnitsTSUnits.GridColumn3tyles.AddRange
 (m UnitscCs)
		Me.dgUnitsTSUnits.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me
		Me
		Me
 .dgUnitsTSUnits.MappingName = "Units"
 .dgUnitsTSUnits.RowHeadersVisible
 .dgUnitsTSUnits.ReadOnly
 End Sub
 = True
 = False
 {Me.
 ()
 ()
 {Me.
 64
 {Me.
 {Me.
 #End Region
 #Region "UnitType"
 Friend WithEvents dgUnitCountsTSUnitType As System.Windows.Forms.DataGridTableStyle
 Friend WithEvents dgUnitCountsTSUnitNationTotals
 DataGridTableStyle
 As System.Windows.Forms.
 Friend WithEvents dgUnitCountsTSUnits As System.Windows.Forms.DataGridTableStyle
 Private
 Sub SetupUnitCountsDGTS ()
 " dgUnitCounts
 ’
 dgUnitCountsTSUnitType (UnitType)
		Me.dgUnitCountsTSUnitType
		Me.dgUnitCounts.TableStyles.AddRange
 = New System.Windows.Forms.DataGridTableStyle()
 {Me.dgUnitCountsTSUnitType})
		Me.dgUnitCountsTSUnitType.DataGrid
 (New System.Windows.Forms.DataGridTableStyle()
 = Me.dgUnitCounts
		Me.dgUnitCountsTSUnitType.GridColumnStyles.AddRange
		Me.dgUnitCountsTSUnitType.HeaderForeColor
 (m UnitTypeCS)
 = System.Drawing.SystemColors.ControlText
		Me.dgUnitCountsTSUnitType.MappingName = "UnitType"
		Me.dgUnitCountsTSUnitType.ReadOnly = True
 '
 dgUnitCountsTSUnitNationTotals
 UnitNationTotals)
		Me.dgUnitCountsTSUnitNationTotals
 Relation
 Nations
 [m _dr3] UnitType to (
 = New System.Windows.Forms.DataGridTableStyle()
		Me.dgUnitCounts.TableStyles.AddRange
 {dgUnitCountsTSUnitNationTotals})
 (New System.Windows.Forms.DataGridTableStyle()
		Me.dgUnitCountsTSUnitNationTotals.DataGrid
 = Me.dgUnitCounts
		Me.dgUnitCountsTSUnitNationTotals.GridColumnStyles.AddRange
		Me.dgUnitCountsTSUnitNationTotals.HeaderForeColor
 (mUnitNationTotalsCS)
 = System.Drawing.SystemColors.
 ControlText
		Me.dgUnitCountsTSUnitNationTotals.MappingName
		Me.dgUnitCountsTSUnitNationTotals.ReadOnly
 = "UnitNationTotals"
 = True
		Me.dgUnitCountsTSUnitNationTotals.RowHeadersVisible
 '
 Units)
 '
 dgUnitCountsTsSUnits
 (AND...)
		Me.dgUnitCountsTS8Units
 Relation
 Relation
 UnitsForTypeAndNation
 UnitsForType
 {m dr7]
 = True
 [m dr8]
 UnitNationTotals
 UnitType to (Units)
 = New System.Windows.Forms.DataGridTableStyle
		Me.dgUnitCounts.TableStyles.AddRange
 {Me .dgUnitCountsTSUnits})
		Me.dgUnitCountsTSUnits.DataGrid
 (New System.Windows.Forms.DataGridTableStyle()
 = Me.dgUnitCounts
		Me.dgUnitCountsTSUnits.GridColumnStyles.AddRange
		Me.dgUnitCountsTSUnits.HeaderForeColor
 (m UnitsCSs)
 ()
 = System.Drawing.SystemColors.ControlText
		Me.dgUnitCountsTSUnits.MappingName = "Units"
		Me.dgUnitCountsTSUnits.RowHeadersVisible
		Me.dgUnitCountsTSUnits.ReadOnly = True
 End Sub
 #End Region
 #Region "Wonders"
 = False
 Friend WithEvents dgWondersTSWonders As System.Windows.Forms.DataGridTableStyle
 Private
 '
 f
 Sub SetupWondersDGTS ()
 dgWonders
 dgWondersT3Wonders (Wonders)
		Me.dgWondersTSWonders
 = New System.Windows.Forms.DataGridTableStyle()
		Me.dgWonders.TableStyles.AddRange
 {Me .dgWondersTSWonders})
		Me.dgWondersTSWonders.DataGrid
 (New System.Windows.Forms.DataGridTableStyle
 = Me.dgWonders
		Me.dgWondersTSWonders.GridColumnStyles.AddRange
 (m WondersCS)
		Me.dgWondersTSWonders.HeaderForeColor = System.Drawing.SystemColors.ControlText
		Me.dgWondersTSWonders.MappingName
		Me.dgWondersTSWonders.ReadOnly
 = "Wonders"
 = True
		Me. dgWondersTSWonders . RowHeadersVisible = False
 End Sub
 #End Region
 #Region "Triumphs"
 Friend WithEvents dgTriumphsTSTriumphs As System.Windows.Forms.DataGridTableStyle
 Private
 Sub SetupTriumphsDGTS ()
 " dgTriumphs
 :
 dgTriumphsTSTriumphs
		Me.dgTriumphsTSTriumphs
 (Triumph)
 = New System.Windows.Forms.DataGridTableStyle
 65
 v4
 4
 4
 v
 to
 ()
 («¢
 «
 ()
		Me.dgTriumphs.TableStyles.AddRange
 .dgTriumphsTSTriumphs})
 (New System.Windows.Forms.DataGridTableStyle
		Me.dgTriumphsTSTriumphs.DataGrid
 = Me.dgTriumphs
		Me.dgTriumphsTSTriumphs.GridColumnStyles.AddRange
		Me.dgTriumphsTSTriumphs.HeaderForeColor
 (m TriumphsCS)
 = System.Drawing.SystemColors.ControlText
		Me.dgTriumphsTSTriumphs.MappingName
 = "Triumph"
		Me.dgTriumphsTSTriumphs.ReadOnly = True
		Me.dgTriumphsTSTriumphs.RowHeadersVisible = False
 End Sub
 #End Region
 #End Region
 #End Region
 #Region "Form Load and Close Routines"
 Private
 Sub Forml Load (ByVal sender As System.Object,
 Handles MyBase.Load
 m Civ
 = New CivLayout
 (Me, progress)
 lblVersion.Text = m
 Civ.GetVersion()
 If
 m
 Civ.IsConvertibleFormat
 Then
 lblTurnsPassed.Text = m
 Civ.TurnsPassed
 lblYear.Text = m
 Civ.Year .
 ByVal e As System.EventArgs)
 lblDifficultylLevel.Text = m
 Civ.GetDifficultylLevel
 lblBarbarianActivity.Text = m
 Civ.GetBarbarianActivity
 lblTurnsOfPeace.Text = m
 Civ.TurnsOfPeace
 lblNumberOfCities.Text = m
 Civ.TotalCities
 lblNumberOfUnits.Text = m
 Civ.TotalUnits
 lblMapWidth.Text
 = m Civ.MapWidth
 1lblMapHeight.Text = m
 Civ.MapHeight
 lblMapArea.Text = m
 Civ.MapArea()
 lblCursorHoriz.Text = m
 Civ.CursorHorizCoord
 llblMapSeed.Text = m
 Civ.MapSeed
 lblCursorVert.Text = m
 Civ.CursorVertCoord
 FillCities()
 FillUnits
 FillNations
 ()
 ()
 FillCivsAndTreaties
 FillWonders
 FillTriumphs
 ()
 ()
 FillUnitTypes()
 FillUnitNationTotals()
 UpdateUnitTypeCounts
 UpdateUnitNation
 FillMapCells
 mSupplyFilter
 mDemandFilter
 m ProducingFilter
 m UnitTypeFilter
 m VetFilter
 ()
 = ""
 FillCmbPlayerColour
 FillCmbCity
 ()
 dvNations.RowFilter
 CheckExcel
 ()
 ()
 ()
 = ""
 = ""
 ()
 = ""
 = nn
 ()
 = "NationHasBeenActive = True”
 progress.Text = "Ready..."
 Else
 End
 End Sub
 Private
 progress.Text = "Incompatible file format...."
 If
 Sub FillCities({)
 progress.Text = "Filling screen with Cities....”
 dsCiv.Tables ("Cities") .Clear()
 Dim 1 As Integer
 For 1 = 0 To m
 _Civ.TotalCities
 Dim r As DataRow = dsCiv.Tables
 Dim c As City = m
 Civ.GetCity(1)- 1
 r ("CityNumber™) = c.CityNumber
 r("CityName") = c.CityName
 r ("VertCoord") = c.VertCoord
 ("Cities")
 .NewRow ()
 66
 ()
 {Me wv
 r ("HorizCoord") = c.HorizCoord
 r ("OwnerColour") = c.OwnerColour.ToString/()
 r("Citysize") = c.CitySize
 r("OrigColour") = c.OrigOwnerColour.ToString()
 r ("WorkingCitySquaresCount")
 r("Palace") = c.Palace
 r ("Barracks") = c.Barracks
 r ("Granary") = c.Granary
 r("Temple") = c.Temple
 = c.WorkingCitySquaresCount
 r ("Marketplace") = c.Marketplace
 r ("Library") = c.Library
 r ("Courthouse") = c.Courthouse
 r("Citywalls") = c.CityWalls
 r ("Aqueduct") = c.Aqueduct
 r ("Bank") = c.Bank
 r ("Cathedral") = c.Cathedral
 r ("University") = c.University
 r("MassTransit") = c.MassTransit
 r{"Colosseum”™) = c.Colosseum
 r ("Factory") = c.Factory
 r ("ManufacturingPlant”) = c.ManufacturingPlant
 r ("SDIDefense") = c.SDIDefense
 r ("RecyclingCentre”) = c.RecyclingCentre
 r ("PowerPlant") = c.PowerPlant
 r ("HydroPlant") = c.HydroPlant
 r {"NuclearPlant") = c.NuclearPlant
 r ("StockExchange™) = c.StockExchange
 r("SewerSystem") = c.SewerSystem
 r ("Supermarket") = c.Supermarket
 r ("Superhighways") = c.Superhighways
 r{"ResearchLab")
 = c.ResearchLab
 r("SAMMissileBattery")
 = c.SAMMissileBattery
 r {"CoastalFortress") = c.CoastalFortress
 r{"SolarPlant") = c.SolarPlant
 r ("Harbor") = c.Harbor
 r ("OffshorePlatform”) = c.OffshorePlatform
 r ("Airport") = c.Airport
 r("PoliceStation") = c.PoliceStation
 r ("PortFacility") = c.PortFacility
 r("CityProducing") = c.CityProducing
 r ("NumberOfActiveTradeRoutes"”")
 II
 = c.NumberOfActiveTradeRoutes
 r("SuppliedCommodityl™)
 r ("SuppliedCommodity2")
 = c.SuppliedCommodity
 Il
 (0)
 c.SuppliedCommodity (1)
 r ("SuppliedCommodity3") = c.SuppliedCommodity (2)
 r ("DemandedCommodityl")
 r ("DemandedCommodity2")
 r ("DemandedCommodity3")
 'r{"TradedCommodityl”}
 ¢ .DemandedCommodity
 (0)
 c.DemandedCommodity (1)
 Il
 c.DemandedCommodity (2)
 = c.GetTradeRoutes
 Extractor\Forml.vb
 ‘Ended up being garbage
 r ("TradedCommodityl™)
 r ("TradedCommodity2")
 r ("TradedCommodity3")
 r ("TradingCityNumberl”)
 c.TradedCommodity (0)
 = c.TradedCommodity (1)
 = c.TradedCommodity (2)
 = c.TradingCity (0)
 r ("TradingCityNumber2") = c.TradingCity (1)
 r ("TradingCityNumber3”) = c.TradingCity (2)
 r("ElvisCount") = c.ElvisCount
 r("ScientistCount”) = c.ScientistCount
 r({"TaxCollectorCount™) = c.TaxCollectorCount
 r ("FoodInStorage")
 = c.FoodInStorage
 r({"ShieldsConsumedInProduction")
 = c.ShieldsConsumedInProduction
 r("TradeFromCitySquares"”)
 = c.TradeFromCitySquares
 r("ScienceForCity") = c.ScienceForCity
 r("TaxForCity") = c.TaxForCity
 r({"TotalTradeForCity") = c.TotalTradeForCity
 r ("FoodFromCitySquares") = c.FoodFromCitySquares
 r("ShieldsFromCity"™) = c¢.ShieldsFromCity
 r ("HappyCitizensInCity") = c.HappyCitizensInCity
 r ("UnhappyCitizensInCity"”)
 = c.UnhappyCitizensInCity
 dsCiv.Tables ("Cities") .Rows.Add(r)
 Next
 Dim col As DataGridColumnStyle
 i=20
 For Each col In dgCities.TableStyles(0).GridColumnStyles
 Dim column As String = col.MappingName
 l1stCityFields.
 Items .Add (column)
 lstCityFields.SetSelected
 1 +=1
 Next
 tabPages.Refresh()
 End Sub
 Private
 Sub FillUnits({()
 (i,
 True)
 progress.Text = "Filling screen with Units...."
 dsCiv.Tables ("Units") .Clear()
 Dim i As Integer
 For i = 0 To m
 Civ.TotalUnits- 1
 Dim r As DataRow = dsCiv.Tables ("Units") .NewRow ()
 Dim u As Unit = m
 Civ.GetUnit
 (i)
 r ("UnitNumber") = u.UnitNumber
 r("VertCoord"”) = u.VertCoord
 r ("HorizCoord") = u.HorizCoord
 r ("Veteran") = u.Veteran
 r("UnitType")
 = u.UnitType
 r("OwnerColour”")
 = u.OwnerColour.ToString/()
 r("HitPoints") = u.HitPoints
 r ("UnitCommodity") = u.UnitCommodity
 r ("UnitOrders”) = u.UnitOrders.ToString()
 r ("VertGotoCoords")
 r ("HorizGotoCoord")
 u.VertGotoCoord
 u.HorizGotoCoord
 r ("HomeCityNumber”™) = u.HomeCityNumber
 r ("UnitRAbove") = u.UnitAbove
 r{"UnitBelow") = u.UnitBelow
 r ("UnitHomeCityName")
 = u.GetHomeCityName (m Civ)
 r("UnitLocation™) = u.UnitLocation(m Civ)
 r ("UnitNationNumber")
 r ("UnitNation")
 r("UnitNear")
 =-1
 = u.UnitNation(m Civ)
 = u.UnitNear (m Civ)
 dsCiv.Tables ("Units") .Rows.Add(r)
 Next
 Dim col As DataGridColumnStyle
 i=20
 For Each col In dgUnits.TableStyles(0).GridColumnStyles
 Dim column As String = col.MappingName
 l1stUnitFields.Items.Add
 lstUnitFields.SetSelected
 i4+=1
 Next
 tabPages.Refresh()
 End Sub
 Private
 Sub FillNations
 ()
 (column)
 (i,
 True)
 progress.Text = "Filling screen with Nations...."
 dsCiv.Tables ("Nation") .Clear ()
 Dim 1 As Byte
 For'i
 Civ.GetNationalityCoumkeass
 = 0 To m
 Dim r As DataRow = dsCiv.Tables
 Dim n As Nationality
 r ("NationNumber")
 = m
 = 1
 Civ.GetNationality
 r ("Nation") = n.Nation.ToString
 1
 ("Nation")
 .NewRow ()
 (i)
 r ("NationColourNumber") = CByte(n.NationColour)
 r("NationColour”) = n.NationColour.ToString
 r("NationActive") = n.NationActive
 r ("NationHasBeenActive") = n.NationHasBeenActive
 r ("NationExtinct"™) = n.NationExtinct
 dsCiv.Tables
 Next
 End Sub
 Private
 ("Nation") .Rows.Add (r)
 Sub FillCivsAndTreaties()
 progress.Text
 = "Filling
 screen with Civilizations
 and Treaties...."
 dsCiv.Tables ("Civilization") .Clear ()
 dsCiv.Tables ("Treaty") .Clear()
 Dim i As Byte
 For i = 0 To m
 Civ.GetCivsCount- 1
 Dim rl As DataRow = dsCiv.Tables
 Dim c¢ As Civilization
 ("Civilization")
 = mCiv.GetCiv(1i)
 rl ("CivColourNumber") = CByte(c.CivColour)
 .NewRow/()
 rl("CivColour") = c.CivColour.ToString
 rl ("Gender") = c.Gender.ToString
 r1("Gold") = c.Gold
 rl ("Researching") = c.Researching
 rl ("ResearchProgress™) = c.ResearchProgress
 rl ("SciencePercent"”)
 = c.SciencePercent
 rl {("TaxPercent") = c.TaxPercent
 rl ("LuxuryPercent”) = c.LuxuryPercent
 rl ("GovernmentType") = c.GovernmentType.ToString
 rl ("AdvancedFlight")
 = c.AdvancedFlight
 rl ("Alphabet") = c.Alphabet
 rl ("AmphibiousWarfare") = c.AmphibiousWarfare
 rl ("Astronomy") = c.Astronomy
 rl ("AtomicTheory") = c.AtomicTheory
 rl ("Automobile") = c.Automobile
 rl ("Banking") = c.Banking
 rl ("BridgeBuilding”™)
 = c.BridgeBuilding
 rl ("BronzeWorking”) = c.BronzeWorking
 rl {"CeremonialBurial")
 = c¢.CeremonialBurial
 rl ("Chemistry") = c.Chemistry
 r1 ("Chivalry") = c.Chivalry
 rl ("CodeofLaws")
 = c.CodeofLaws
 rl ("CombinedArms"”) = c.CombinedArms
 rl ("Combustion") = c.Combustion
 rl ("Communism") = c.Communism
 rl ("Computers") = c.Computers
 rl ("Conscription™) = c.Conscription
 rl ("Construction") = c.Construction
 rl ("Corporation") = c.Corporation
 rl ("Currency") = c.Currency
 rl ("Democracy")
 = c.Democracy
 rl ("Economics") = c.Economics
 rl ("Electricity") = c.Electricity
 rl("Electronics”) = c.Electronics
 rl ("Engineering") = c.Engineering
 rl ("Environmentalism")
 = c.Environmentalism
 rl ("Espionage") = c.Espionage
 rl ("Explogives™) = c.Explosives
 rl ("Feudalism") = c.Feudalism
 rl "Eisght") = ¢.Flight
 rl ("Fundamentalism") = c.Fundamentalism
 rl ("FusionPower") = c.FusionPower
 rl ("GeneticEngineering") = c.GeneticEngineering
 rl ("GuerillaWarfare”) = c.GuerillaWarfare
 rl ("Gunpowder”™) = c.Gunpowder
 rl ("HorsebackRiding") = c.HorsebackRiding
 rl ("Industrialization") = c.Industrialization
 rl ("Invention") = c.Invention
 rl ("IronWorking") = c.IronWorking
 rl ("LaborUnion") = c.LaborUnion
 rl ("Laser") = c.Laser
 rl ("Leadership") = c.Leadership
 rl("Literacy") = c.Literacy
 rl ("MachineTools") = c.MachineTools
 rl ("Magnetism") = c.Magnetism
 rl ("MapMaking”) = c.MapMaking
 rl ("Masonry") = c.Masonry
 rl ("™MassProduction") = c.MassProduction
 rl ("Mathematics") = c.Mathematics
 rl ("Medecine") = c.Medecine
 rl ("Metallurgy") = c.Metallurgy
 rl("™™Minjiaturization™) = c¢.Miniaturization
 rl ("MobileWarfare”) = c.MobileWarfare
 rl ("Monarchy") = c¢.Monarchy
 rl ("Monotheism") = c.Monotheism
 rl ("Mysticism") = c.Mysticism
 rl ("Navigation") = c.Navigation
 rl ("NuclearFission") = c.NuclearFission
 rl ("NuclearPower"”)
 = c.NuclearPower
 rl ("Philosophy") = c.Philosophy
 rl ("Physics") = c.Physics
 rl ("Plastics") = c.Plastics
 rl ("Plumbing") = c.Plumbing
 69
 rl ("Polytheism") = c.Polytheism
 rl ("Pottery") = c.Pottery
 rl ("Radio") = c.Radio
 rl ("Railroad")
 = c¢.Railroad
 rl ("Recycling”) = c.Recycling
 rl ("Refining") = c.Refining
 rl
 ("Refrigeration™)
 =
 c.Refrigeration
 rl ("Republic") = c.Republic
 rl ("Robotics") = c.Robotics
 rl ("Rocketry") = c.Rocketry
 rl ("Sanitation") = c.Sanitation
 rl("Seafaring”) = c.Seafaring
 rl ("SpaceFlight") = c.SpaceFlight
 rl1 ("Stealth")
 = c.Stealth
 rl ("SteamEngine”) = c.SteamEngine
 r1("Steel”™) = ¢.S8teel
 rl
 ("Superconductor™)
 =
 rl (™fasctics™) = c.Tactics
 c.Superconductor
 rl ("Theology") = c.Theology
 rl
 ("TheoryofGravity")
 =
 c.TheoryOfGravity
 rl ("Trade") = c.Trade
 rl ("University") = c.University
 rl ("WarriorCode"™) = c.WarriorCode
 rl ("Wheel") = c.Wheel
 rl ("Writing") = c.Writing
 dsCiv.Tables
 Dim Jj As Byte
 For
 ("Civilization")
 Jj = 0 To 7
 .Rows.Add (rl)
 Dim pc As CivLayout.PlayerColour
 PlayerColour)
 Dim r2 As DataRow =
 r2 ("FromCivColourNumber")
 dsCiv.Tables
 =
 =
 CType(((]J
 ("Treaty")
 CByte (c.CivColour)
 r2 ("ToCivColour")
 r2 ("ToCivNation")
 =
 pc.ToString
 + 1) Mod 8), CivLayout.
 .NewRow ()
 = Nationality.GetNationalityNameByColour(m
 If
 Else
 7 < 7 Then
 Dim t As Civilization.TreatyStructure
 r2 ("Contact")
 r2 ("CeaseFire")
 r2 ("Peace")
 =
 =
 t.Contact
 = t.CeaseFire
 Tt. Peace
 r2 ("Alliance") = t.Alliance
 r2 ("Vendetta") = t.Vendetta
 r2 ("Embassy") = t.Embassy
 r2 ("war") = t.War
 r2 ("Contact") = False
 r2 ("CeaseFire") = False
 r2 ("Peace") = False
 r2 ("Alliance") = False
 r2 ("Vendetta") = False
 r2 ("Embassy") = False
 r2 ("War") = False
 End If
 r2 ("Attitude") = c.Attitude (J)
 r2 ("LastContactTurn")
 r2 ("LastContactYear")
 =
 =
 =
 c.Treaty(])
 c.LastContactTurn(])
 m Civ.CalcYearforTurn(c.LastContactTurn(j))
 dsCiv.Tables
 Next
 Next
 End Sub
 Private
 Sub Fillwonders()
 progress.Text
 =
 ("Treaty")
 .Rows.Add (r2)
 "Filling screen with Wonders...."
 dsCiv.Tables
 ("Wonders") .Clear()
 Dim 1 As Integer
 For 1 = 0 To m
 Civ.GetWondersCount- 1
 Dim r As DataRow =
 Dim w As Wonder =
 dsCiv.Tables
 ("Wonders")
 Civ.GetWonder (i)
 m
 r {"WonderNumber®)
 r ("WonderName")
 =
 =
 w.WonderNumber
 w.WonderName
 r ("WonderEra")
 w.WonderEra.ToString
 =
 Tf w.WonderCity Is Nothing Then
 r ("WonderCityNumber™) = 255
 r ("WonderCityName")
 r ("WonderCityColour®)
 = ""
 .NewRow ()
 Civ, pc)
 70
 = ""
 Else
 r ("WonderCityNumber")
 = w.WonderCityNumber
 r ("WonderCityName") = w.WonderCityName
 r ("WonderCityColour") = w.WonderCity.OwnerColour.ToString
 End If
 r ("WonderBuilt") = w.WonderBuilt
 r ("WonderDestroyed"”") = w.WonderDestroyed
 dsCiv.Tables
 ("Wonders") .Rows.Add (r)
 Next
 End Sub
 Private
 Sub FillTriumphs ()
 progress.Text = "Filling screen with Triumphs...."
 dsCiv.Tables ("Triumph") .Clear()
 If
 m
 Civ.GetTriumphCount
 Exit
 End If
 Sub
 Dim 1 As Byte
 For i = 0 To m
 = 0 Then
 Civ.GetTriumphCount
 Dim r As DataRow = dsCiv.Tables
 Dim t As Triumph = m
 Civ.GetTriumph
 r ("TriumphNumber")
 = i- 1
 ("Triumph")
 (i)
 r ("TriumphNationNumber") = CByte(t.Nation)
 r ("TriumphNation™) = t.Nation.ToString
 r{"TriumphYear"”)
 = t.TriumphYear
 r("TriumphTurn”) = t.TriumphTurn
 r("TriumphNationColour")
 = m
 .NewRow ()
 Civ.GetNationality(t.Nation).NationColour.ToString
 dsCiv.Tables ("Triumph") .Rows.Add(r)
 Next
 End Sub
 Private
 Sub FillUnitTypes()
 progress.Text = "Filling screen with Unit Types...."
 ReDim m ut(m Civ.m CivUnit.Length
 m Civ.m CivUnit.CopyTo(m
 Array.Sort (m ut)
 ut,
 dsCiv.Tables ("UnitType") .Clear()
 Dim 1 As Byte
 ut.Length- 1
 For 1 = 0 To m
 Dim r As DataRow = dsCiv.Tables
 #{*UnitType”) = m
 uti)
 r("UnitTypeCount"”)
 dsCiv.Tables
 = 0
0)
 1)
 ("UnitType")
 ("UnitType") .Rows.Add (r)
 Next
 End Sub
 Private
 Sub FillUnitNationTotals()
 progress.Text
 .NewRow ()
 = "Updating Unit/Nation Statistics...."
 dsCiv.Tables ("UnitNationTotals") .Clear()
 dvUnitCounts.AllowEdit
 = True
 Dim rl, r2 As DataRow
 Dim c¢ As CivlLayout.PlayerColour
 For Fach rl In dsCiv.Tables
 c = c.Parse(c.GetType,
 ("Units")
 .Rows
 rl.Item{("OwnerColour™))
 Dim n As Nationality
 = mCiv.GetNationalityByColour(c,
 Dim key (l) As Object
 key (0) = rl.Item("UnitType")
 Tf n Is Nothing Then
 key (l) = Nothing
 Else
 key(l) = n.Nation
 End If
 Dim i As Integer
 Tf i € 0 Then
 r2 = dsCiv.Tables
 = dvUnitCounts.Find
 ("UnitNationTotals")
 (key)
 r2 ("UnitType") = rl.Item("UnitType")
 r2 ("NationNumber"™)
 = CByte (key (l))
 If n Is Nothing Then
 r2 ("Nation")
 = nun
 r2 ("UnitNationColour®™)
 Else
 = ""
 r2 ("Nation") = n.Nation.ToString
 .NewRow ()
 True)
 71
 r2 ("UnitNationColour") = n.NationColour.ToString
 End If
 r2 ("UnitNationCount")
 dsCiv.Tables
 = 1
 ("UnitNationTotals")
 Else
 End
 Next
 dvUnitCounts.Item(i).Item("UnitNationCount")
 If
 dvUnitCounts.AllowEdit
 End Sub
 Private
 = False
 Sub UpdateUnitTypeCounts
 progress.Text
 ()
 .Rows.Add (r2)
 = "Updating Unit Type Statistics....”
 dvUnitTypes.AllowEdit = True
 Dim FirstTime
 As Boolean = True
 Dim c¢ As Short = 0
 Dim 1 As Integer
 Dim LastUnitType
 As String
 For 1 = 0 To dvUnitCounts.Count- 1
 Dim r As DataRowView
 r = dvUnitCounts.Item(1i)
 If
 FirstTime
 = True Then
 FirstTime = False
 LastUnitType = r.Item("UnitType")
 ElseIf
 r.Item("UnitType")
 <> LastUnitType Then
 UpdateUnitType(c, LastUnitType)
 LastUnitType = r.Item("UnitType")
 c =0
 End If
 c += r.Item("UnitNationCount")
 Next
 If
 FirstTime
 = False Then
 UpdateUnitType(c, LastUnitType)
 End If
 dvUnitTypes.AllowEdit
 dvUnitTypes.RowFilter
 End Sub
 Private
 False
 "UnitTypeCount
 <> 0"
 += 1
 Sub UpdateUnitType (ByVal c As Short, ByVal u As String)
 Dim i As Integer
 If
 i
 >= 0 Then
 = dvUnitTypes.Find(u)
 dvUnitTypes.Item (i) .BeginEdit ()
 dvUnitTypes.Item (i) .Item("UnitTypeCount"”) = c
 dvUnitTypes.Item (i) .EndEdit ()
 End
 End Sub
 Private
 If
 Sub UpdateUnitNation()
 Dim i As Integer
 For i = 0 To dsCiv.Units.Count- 1
 Dim r As DataRow
 r = dsCiv.Units.Item(i)
 Dim c¢ As CivLayout.PlayerColour
 Dim n As Nationality
 c = c.Parse(c.GetType, r.Item("OwnerColour"))
 n = m
 Civ.GetNationalityByColour(c,
 If n Is Nothing Then
 dsCiv.Units.Item(i).Ttem("UnitNationNumber")
 dsCiv.Units.Ttem(i).Item("UnitNation™)
 Else
 True)
 dsCiv.Units.Item(i).Item("UnitNationNumber"™)
 dsCiv.Units.Item(i).Item("UnitNation”)
 End IL
 Next
 End Sub
 Private
 Sub FillMapCells/()
 progress.Text
 = "Filling
 screen
 OO
 = ""
 = CByte (n.Nation)
 = n.Nation.ToString
 with Map Cell References....”
 dsCiv.Tables
 Dim 1 As Short
 ("MapCell™) .Clear()
 For 1 = 0 To m
 Civ.MapArea —-1
 Dim mc As MapCell = m
 Civ.GetMapCell
 If me.CityOrUnitIsPresent
 Dim r As DataRow = dsCiv.Tables
 (i)
 = True Then
 ("MapCell"™)
 72
 .NewRow ()
 r ("MapCellNumber")
 r ("MapHorizCoord")
 r ("MapVertCoord")
 = i
 = CShort (mc.Width)
 = CShort (mc.Height)
 r ("MapCivNation") = mc.GetUnitNationName (m Civ)
 r("MapCivColour") = mc.GetUnitColour
 r ("MapCityName") = mc.GetWorkedByCityName
 r ("MapCellNear"”) = mc.GetNear
 r ("MapNumberOfUnits")
 = mc.GetNumberofUnits
 r ("MapDescOfUnits") = mc.GetUnits(m Civ, False)
 dsCiv.Tables ("MapCell") .Rows.Add(r)
 End If
 Next
 End Sub
 Private
 Sub FillCmbPlayerColour()
 progress.Text
 = "Filling Player Colours Combobox...."
 Dim ¢ As CivLayout.PlayerColour
 Dim t As Type = c.GetType
 Dim s() As String
 s = ¢.GetNames (t)
 cmbOwnerColour.Items.Add ("All")
 cmbOwnerColour.Items.AddRange(s)
 BuildFilter
 = True
 m
 cmbOwnerColour.SelectedIndex
 End Sub
 Private
 Sub FillCmbCity()
 progress.Text
 = 0
 = "Filling Cities Combobox..."
 CmbCity.BeginUpdate
 CmbCity.Items.Clear
 ()
 ()
 CmbCity.Items.Add ("All")
 Dim ct As City
 For Each ct In mCiv.GetCities
 CmbCity.Items.Add(ct)
 Next
 CmbCity.EndUpdate()
 mBuildFilter = False
 CmbCity.SelectedItem = "All"
 m
 End Sub
 BuildFilter
 Private
 = True
 Sub CheckExcel ()
 Dim i As Integer
 x1lfname = m
 = m
 Civ.Filename.IndexOf(".")
 Civ.Filename.Substring(0,
 If System.IO.File.Exists({xlfname)
 btnExcel.Enabled = False
 Else btnExcel.Enabled
 End
 End Sub
 Private
 If
 i
 + 1) & "xls"
 = True Then
 = True
 Sub Forml Closing (ByVal sender As Object, ByVal e As System.ComponentModel.
 CancelEventArgs) Handles MyBase.Closing
 tmrSecond.Stop()
 End Sub
 #End Region
 #Region "Tab Pages, Buttons and Datagrid Navigation on Relationships”
 Private
 Sub btnExcel Click(ByVal sender As System.Object,
 Handles btnExcel.Click
 btnExcel.Enabled = False
 ExcelMap ()
 progress.Text = "Ready..."
 End Sub
 #Region
 "Excel
 Private
 Map Extraction”
 Sub ExcelMap ()
 ExcelProgressBarSetup()
 Dim x1App As Excel.Application
 = CType (CreateObject
 ByVal e As System.EventArgs)
 ("Excel.Application"),
 Excel.
 73
 Application)
 Dim x1Book As Excel.Workbook
 Dim x1Sheet As Excel.Worksheet
 Studio «Projects\CIV = CType (xlApp.Workbooks.Add,
 = CType (x1Book.Worksheets(l),
 ExcelColumnHeadings (x1Sheet)
 Dim 1 As Short
 For i = 0 To m
 Civ.MapArea- 1
 ExcelAddRow (x1Sheet, i)
 Next
 ExcelFormat (x1Sh
 ExcelSave (x1Shee
 End Sub
 Private
 eet,
 t,
 x1App)
 x1Book,
 Sub ExcelProgressBarSetup()
 x1App)
 * Set text in StatusBarPanel
 progress.Text
 '
 =
 "Creating
 Excel
 Setup Progressbar information
 pbExcel . Minimum
 pbExcel .Maximum
 pbExcel.Value
 pbExcel.Step
 pbExcel.Visible
 '
 =
 = 1
 = 0
 = mCiv.MapArea
 0
 = True
 Make Excel Progress
 pnlExcel.Visible
 pnlExcel.Refresh()
 End Sub
 Private
 = True
 panel
 Spreadsheet
visible
 1
 of Map Data..."
 Sub ExcelColumnHeadings (ByRef xl1Sheet As Excel.Worksheet)
 x1Sheet.Cells (1,
 x1Sheet.Cells (1,
 x1Sheet.Cells (1,
 Xx1lSheet.Cells (1,
 x1Sheet.Cells (1,
 x1Sheet.Cells (1,
 xlSheet.Cells (1,
 x1Sheet.Cells (1,
 Xx1lSheet.Cells (1,
 x1Sheet.Cells (1,
 XlSheet.Cells (1,
 xXx1Sheet.Cells (1,
 x1Sheet.Cells (1,
 X1lSheet.Cells (1,
 X1Sheet.Cells (1,
 x1Sheet.Cells (1,
 Xx1Sheet.Cells (1,
 X1lSheet.Cells (1,
 x1lSheet.Cells (1,
 x1Sheet.Cells (1,
 x1Sheet.Cells (1,
 X1Sheet.Cells (1,
 x1Sheet.Cells (1,
 x1Sheet.Cells (1,
 xlSheet.Cells
 (1,
 X1lSheet.Cells (1,
 xX1lSheet.Cells (1,
 x18heet.Cells (1,
 xlsheet.Cells (1,
 XlSheet.Cells (1,
 x1Sheet.Cells (1,
 x18heet.Cells (1,
 End Sub
 Private
 1)
 2)
 3)
 4)
 5)
 0)
 7)
 8)
 9)
 21)
 22)
 23)
 24)
 25)
 26)
 27)
 28)
 29)
 30)
 31)
 32)
 "Width"
 "Height"
 "Address"
 — i
 a
 "TT Desc"
 ps
 nen
 LL ~ 44
 "Riv"
 134 Far
 PAE"
 14] Pol
 Im
 "For"
 —_ "Rai"”
 "Roa
 = "Min"
 Tr Tre”
 T¥
 wo Ew”
 "Unt"
 "Cont#"
 144
 Tr
 "Owned Col"
 = "Near"
 = "Worked By"
 eg
 "Units"
 ad
 Li}
 ak”
 a
 Lf
 HRY
 w G"
 =
 "RE"
 np
 TY
 Sub ExcelAddRow (ByRef x1Sheet As Excel.Worksheet,
 Excel.Workbook)
 Excel.Worksheet)
 ByVal i As Short)
 "Update StatusBarPanel
 progress .Text = "Creating
 of
 " &m Civ.MapArea.ToString
 '
 Add a row of data..
 Dim mc As MapCell = m
 with current
 Excel
 row statistics.
 Spreadsheet
 fe “leans”
 m Civ.
 X1Sheet.Cells(i
 GetMapCell (i)
 + 2, 1) = mc.Width
 x1lSheet.Cells
 X1Sheet.Cells
 (1
 (i
 + 2,
 + 2,
 + 2,
 2
 3
 5
 ) = mc.Height
 of Map Data.
 ) = mCiv.GetMapDataBlock2BaseAddress
 X1Sheet.Cells
 (1
 )
 If mc.River = True Then
 X1Sheet.Cells(i
 = mc.TerrainBaseType.ToString
 + 2, 8)
 x18heet.Cells
 (1
 + 2, 4)
 1
 mc.TerrainBaseType
 EK"
 + 128
 & (1 + 1).ToString
 + i
 74
 & "
 "4
 Else
 x1Sheet.Cells(i
 + 2, 4) = mc.TerrainBaseType
 End If
 If mc.ExtraShield
 x1Sheet.Cells(i
 Bae
 LT
 = True Then
 + 2, 6) = 1
 If mc.PotentialExtraShield
 xX1lSheet.Cells(i
 = True Then
 + 2, 7) = 1
 Bmel IL
 If mc.Farmland Then
 x1Sheet.Cells(i1
 End If
 If mc.Airbase Then
 x1lSheet.Cells
 End If
 If mc.Pollution
 (i
 + 2, 9) = 1
 + 2, 10)
 = True Then
 x1lsSheet.Cells(i
 it
 =
 + 2, 11) = 1
 End
 If
 If mc.Fortress
 x1Sheet.Cells(i
 End
 ET
 If mc.Railroad
 x1S8heet.Cells
 Engl It
 = True Then
 + 2, 12) = 1
 = True Then
 (i
 + 2, 13) = 1
 If mc.Road = True Then
 x1Sheet.Cells(i
 End
 If
 If mc.Mined Then
 X1lSheet.Cells(i
 End
 If
 If mc.Irrigated
 X1Sheet.Cells(1
 End
 If
 If mc.CityPresent
 x1Sheet.Cells(i
 + 2, 14) = 1
 + 2, 15) = 1
 Then
 + 2, 16)
 = True Then
 I
 —
 + 2, 17) = 1
 x1lSheet.Cells(i
 End If
 If mc.UnitPresent
 X1lSheet.Cells(i1
 End IT
 X1Sheet.Cells
 + 2, 23) = mc.WorkedBy.CityName.ToString
 = True Then
 + 2, 18) = 1
 (i + 2, 19) = mc.ContinentNumber
 If Not (mc.OwnedColour = MapCell.OwnedColourEnum.NoneQOrBarbarian)
 X1Sheet.Cells
 End If
 x1lSheet.Cells
 (i + 2, 20) = mc.OwnedColour.ToString
 (i + 2, 21) = mc.GetNear
 If Not (mc.WorkedBy Is Nothing) Then
 x1lSheet.Cells(i
 End If
 x1Sheet.Cells(i
 If mc.VisibleToPurple
 Xx1lSheet.Cells(i
 Ered ~T 6
 If mc.VisibleToOrange
 x1Sheet.Cells(i
 End If
 + 2, 22) = mc.WorkedBy.CityName.ToString
 + 2, 24) = mc.GetUnits(m Civ, True)
 = True Then
 + 2, 25) = 1
 = True Then
 + 2, 26) = 1
 If mc.VisibleToTurquolse
 x1Sheet.Cells
 (i
 = True Then
 + 2, 27) = 1
 End If
 If mc.VisibleToYellow
 x18heet.Cells
 End If
 If mec.VisibleToBlue
 x1Sheet.Cells
 (i
 = True Then
 + 2, 28) = 1
 = True Then
 (1 + 2, 29) = 1
 End If
 If mc.VisibleToGreen
 x1Sheet.Cells
 End If
 = True Then
 (i
 + 2, 30) = 1
 If mc.VisibleToWwhite = True Then
 xX1Sheet.Cells (1 + 2, 31) = 1
 End If
 If mc.VisibleToRed
 xiSheet . Cells(i
 = True Then
 + 2, 32) = 1
 re
 Then
 75
 Endl JE
 (es
 ¥
 fo KF oFMa
 P°
 =
 NT
 tJ
 Oo
 5 Ta ET Pa)
 TY em aon
 C808
 Py
 Y
 £9
 i
 ry
 NF at nd A <0 1 11 ERR RT
 he
 2 bd Trae —
 pbExcel. PerformStep ()
 End Sub
 Private
 Sub ExcelFormat
 ex ¥
 om
 ose
 (ByRef xl1lSheet As Excel.Worksheet,
 ByRef xlapp As Excel.Application)
 progress.Text
 = Bc
 Map Data Excel Spreadsheet
 Layout...."
 This inform
 xithset
 ation
Haplieas ion, Visible = True
 x1Sheet.Cells.Select()
 xlapp.Selection.autofilter
 xlapp.Selection.AutoFilter
 x1Sheet.Cells.EntireColumn.AutoFit
 x1Sheet.Rows ("2:2") .select()
 xlapp.ActiveWindow.FreezePanes
 End Sub
 Private
 Sub ExcelSave (ByRef xl1Sheet As Excel.Worksheet,
 ByRef xl1Book As Excel.Workbook,
 ByRef xlapp As Excel.Application)
 Dim e As Boolean = False
 progress.Text
 Do
 LEY
 = "Saving
 If
 System.IO.File.Exists(xlfname)
 System.IO.File. Delete (xlfname)
 End If
 xX1Sheet.SavelAs (x1 fname)
 Catch ex As Exception
 MsgBox (ex.Message.ToString)
 = True
 End Try
 Loop Until e = False
 x1lBook.Close()
 xlapp.Quit
 ()
 x1Sheet = Nothing
 x1Book = Nothing
 xlapp = Nothing
 '
 Hide Excel progress panel
 pnlExcel.Visible = False
 MsgBox ("Saved
 End Sub
 #End Region
 Private
 Map Data Excel Spreadsheet
 Sub BtnFillTerrain
 ByVal
 Setup
 ¥
 Pa
 e = System. EventArgs)
 IEE
 ile
 Fr Ge
 Dim nfn Re New StringBuilder
 Dim i As Integer
 nfn.Append (m Civ.Filename.Substring(0,
 Dim nf As String = m Civ.Filename.Substring(i
 en
 wo
 po
 oy om go
 1
 j=
 JN
 was gleaned from Excel macro generator/editor
 LEELA fs whe BS
 : 9
 ()
 (Field:=21,
 pot
 Criterial:="<>")
 ()
 = True
 Map Data Excel Spreadsheet
 Layout
 = True Then
 Layout
 as
 :
 )
 '
 Near
 as
 " & xlfname)
 Click (ByVal sender As System.Object,
 Handles
 (Place
 doe
 d= ~~ wr
 for dd
 Sr
 Ww
 oi
 §
 ot
 ()
 FHL Jer Sear A Click
 for
 24
 = mCiv.Filename.LastIndexOf
 Plains filled terrain
 ("\")
 i + 1))
 dw &
 :
 " & xlfname
 in place of
 + 1, m Civ.Filename.Length
 nfn.Append (nf. Substring (0, 2))
 nfn.Append ("P")
 nfn.Append (nf.Substring
 '
 (3, nf.Length
 Purge file 1f it already exists
 If File.Exists(nfn.ToString)
 File.Delete (nfn.ToString)
 End IL
 Dim ¢{m Civ.b.Length)
 m Civ.b.CopyTo(c,
 As Byte
 0)
 Then
 For 1 = 0 To m
 Civ.MapArea- 1
 c(m Civ.GetMapDataBlockZBaseAddress
 Next
 Dim fs As FileStream- 3))
 + (i * 6)) = 129
 = New FileStream(nfn.ToString,
 Dim w As BinaryWriter
 = New BinaryWriter
 For i = 0 To oc.ength —1
 w.Write(c(i))
 Next
 w.Flush
 Fo.Flush()
 w.Close()
 fs.Close()
 ()
 (fs)
 FileMode.CreateNew)
 [- iid
 1)
 End Sub
 Private
 Sub tabColSelect
 77
 Layout (ByVal sender As Object, ByVal e As System.Windows.Forms. «
 LayoutEventArgs) Handles tabColSelect.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabNat Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabNat.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabCiv Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabCiv.Layout
 llblOwnerColour.Visible
 = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 I
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabCities
 Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabCities.Layout
 llblOwnerColour.Visible = True
 cmbOwnerColour.Visible = True
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 grpCounters.Visible
 grpCounters.Text
 = True
 = False
 = True
 = "City Counts”
 Dim i As Integer
		Me.dgCitiesTSCities.GridColumnStyles.Clear
 For i = 0 To lstCityFields.Items.Count
 If
 lstCityFields.SelectedIndices.Contains
 Select
 Case lstCityFields.Items.Item(i).ToString/()
 Case "CityNumber"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "CityName"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "VertCoord"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "HorizCoord"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "OwnerColour"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "CitySize"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "OrigColour"”
		Me.dgCitiesT3Cities.GridColumnStyles.Add
 Case "WorkingCitySquaresCount"
		Me.dgCitlesTSCities.GridColumnStyles.Add
 Case "Palace"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Barracks"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Granary"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Temple"
		Me.dgCitiesTSCities.GridColumnsStyles.Add
 Case "Marketplace"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Library"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Courthouse"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 4
 v4
 v4
 ()- 1
 (i)
 Then
 (Me.CityNumber)
 (Me.CityName)
 (Me.CityVertCoord)
 (Me.CityHorizCoord)
 (Me.CityOwnerColour)
 (Me.CitySize)
 (Me.OrigColour)
 (Me.WorkingCitySquaresCount)
 (Me.Palace)
 (Me.Barracks)
 (Me.Granary)
 (Me.Temple)
 (Me.Marketplace)
 (Me.Library)
 (Me.Courthouse)
 Case "CityWalls™
		Me.dgCitiesTSCities
 .GridColumnStyles.Add (Me.CityWalls)
 Case
 Case
 "Aqueduct"
		Me.dgCitiesTSCities
 "Bank"
 .GridColumnStyles.Add
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Cage "Cathedral"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "University"
 (Me.Aqueduct)
 (Me.Bank)
 (Me.Cathedral)
		Me.dgCitiesTSCities.GridColumnStyles.Add (Me.University)
 Case "MassTransit"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Colosseum"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Factory"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "ManufacturingPlant"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case
 "SDIDefense"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "RecyclingCentre"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case
 "PowerPlant"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "HydroPlant"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "NuclearPlant"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "StockExchange"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "SewerSystem"
		Me.dgCitiesTsSCities.GridColumnStyles.Add
 Case "Supermarket”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Superhighways"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case
 Case
 case
 "ResearchLab"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 "SAMMissileBattery"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 "CoastalFortress"”
 (Me.MassTransit)
 (Me.Colosseum)
 (Me.Factory)
 (Me.ManufacturingPlant)
 (Me.SDIDefense)
 (Me.RecyclingCentre)
 (Me.PowerPlant)
 (Me.HydroPlant)
 (Me.NuclearPlant)
 (Me.StockExchange)
 (Me.SewerSystem)
 (Me. Supermarket)
 (Me.Superhighways)
 (Me.ResearchLab)
 (Me.SAMMissileBattery)
		Me.dgCitiesTSCities.GridColumnStyles.Add (Me.CoastalFortress)
 Case
 Case
 Case
 "SolarPlant"
		Me.dgCitiesTSCities.GridColumnStyles.Add (Me.SolarPlant)
 "Harbor"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 "OffshorePlatform"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "Airport"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case
 "PoliceStation"”
 (Me.Harbor)
 (Me.OffshorePlatform)
 (Me.Airport)
		Me.dgCitiesTSCities.GridColumnStyles.Add (Me.PoliceStation)
 Case "PortFacility"
		Me.dgCitiesTSCities.GridColumnStyles.Add(Me.PortFacility)
 Case "CityProducing"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case
 NumberofActiveTradeRoutes)
 "NumberOfActiveTradeRoutes"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "SuppliedCommodityl"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "SuppliedCommodity2"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "SuppliedCommodity3"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "DemandedCommodityl"™
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "DemandedCommodity2"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "DemandedCommodity3"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradedCommodityl"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradedCommodity2"
 (Me.CityProducing)
 (Me.
 (Me. SuppliedCommodityl)
 (Me.SuppliedCommodity?2)
 (Me. SuppliedCommodity3)
 (Me.DemandedCommodityl)
 (Me.DemandedCommodity2)
 (Me.DemandedCommodity3)
 (Me.TradedCommodityl)
 78
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradedCommodity3"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradingCityNumberl"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradingCityNumber2"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TradingCityNumber3"
		Me.dgCitiesgTSCities.GridColumnStyles.Add
 Case "ElvisCount”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "ScientistCount”
 (Me.TradedCommodity2)
 (Me.TradedCommodity3)
 (Me.TradingCityNumberl)
 (Me.TradingCityNumber2)
 (Me.TradingCityNumber3)
 (Me.ElvisCount)
		Me.dgCitiesTSCities.GridColumnStyles.Add (Me.ScientistCount)
 Case "TaxCollectorCount"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "FoodInStorage"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "ShieldsConsumedInProduction"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 ShieldsConsumedInProduction)
 Case "TradeFromCitySquares"”
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "ScienceForCity"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "TaxForCity"
		Me.dgCitiesTSCitilies.GridColumnStyles.Add
 Case "TotalTradeForCity"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "FoodFromCitySquares"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "ShieldsFromCity"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "HappyCitizensInCity"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 Case "UnhappyCitizensInCity"
		Me.dgCitiesTSCities.GridColumnStyles.Add
 End Select
 End If
 Next
 FillCitiesPnlCount
 SetupUnitFilter
 End Sub
 Private
 ()
 ()
 (Me.TaxCollectorCount)
 (Me.FoodInStorage)
 (Me.
 (Me.TradeFromCitySquares)
 (Me.ScienceForCity)
 (Me.TaxForCity)
 (Me.TotalTradeForCity)
 (Me.FoodFromCitySquares)
 (Me.ShieldsFromCity)
 (Me.HappyCitizensInCity)
 (Me.UnhappyCitizensInCity)
 Sub tabUnits Layout (ByVal sender As Object, ByVal e€As System.Windows,Forms.
 LayoutEventArgs) Handles tabUnits.lLayout
 lblOwnerColour.Visible
 = True
 cmbOwnerColour.Visible = True
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 grpCounters.Visible
 = False
 = True
 = True
 CmbCity.Enabled = True
 grpCounters.Text
 Dim 1 As Integer
 = "Unit Counts”
		Me.dgUnitsTSUnits.GridColumnStyles.Clear
 For i = 0 To 1lstUnitFields.Items.Count
 If
 lstUnitFields.SelectedIndices.Contains(i)
 Select
 ()- 1
 Case lstUnitFields.Items.Item(i).ToString()
 Case "UnitNumber"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "VertCoord"
 Then
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "HorizCoord"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "Veteran"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitType"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "OwnerColour™
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "HitPoints"™
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitCommodity"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 (Me.UnitNumber)
 (Me.UnitVertCoord)
 (Me.UnitHorizCoord)
 (Me.Veteran)
 (Me.UnitType)
 (Me.UnitOwnerColour)
 (Me.HitPoints)
 (Me. UnitCommodity)
 Case "UnitOrders"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "VertGotoCoords™”
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "HorizGotoCoord"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "HomeCityNumber"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitAbove"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitBelow"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitHomeCityName"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitLocation"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitNationNumber"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitNation"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 Case "UnitNear"
		Me.dgUnitsTSUnits.GridColumnStyles.Add
 End Select
 End
 If
 Next
 FillUnitsPnlCount
 End Sub
 Private
 ()
 Sub tabUnitCount
 (Me.UnitOrders)
 (Me.UnitVertGotoCoords)
 (Me.UnitHorizGotoCoords)
 (Me.HomeCityNumber)
 (Me.UnitAbove)
 (Me.UnitBelow)
 (Me.UnitHomeCityName)
 (Me.UnitLocation)
 (Me.UnitNationNumber)
 (Me.UnitNation)
 (Me.UnitNear)
 80
 Layout (ByVal sender As Object, ByVal e As System.Windows.Forms. «
 LayoutEventArgs) Handles tabUnitCount.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabMapCells Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabMapCells.Layout
 lblOwnerColour.Visible
 = True
 cmbOwnerColour.Visible = True
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabWonders Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabWonders.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabTriumphs Layout (ByVal sender As Object, ByVal e As System.Windows.Forms.
 LayoutEventArgs) Handles tabTriumphs.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 End Sub
 Private
 Sub tabSummary Layout (ByVal sender As Object,
 LayoutEventArgs) Handles tabSummary.Layout
 lblOwnerColour.Visible = False
 cmbOwnerColour.Visible = False
 pnlCityFilter.Visible
 pnlUnitFilter.Visible
 = False
 = False
 grpCounters.Visible = False
 «
 «
 «
 ByVal e As System.Windows.Forms.
 «
 End Sub
 Private
 Sub FillCitiesPnlCount
 ()
 count.Clear (count, 0, count.Length)
 Dim r As DataRow
 For Each r In dsCiv.Tables
 ("Cities")
 .Rows
 Select
 Case r.Item("OwnerColour")
 Case "White"
 count (CivLayout.PlayerColour.White)
 Case "Green"
 count (CivLayout.PlayerColour.Green)
 Case "Blue"
 count (CivLayout.PlayerColour.Blue)
 Case "Yellow"
 count (CivLayout.PlayerColour.Yellow)
 Case "Turquoise"
 count (CivLayout.PlayerColour.Turquoise)
 Case "Orange"
 count (CivLayout.PlayerColour.Orange)
 Case "Purple"
 count (CivLayout.PlayerColour.
 Case "Red"
 count (CivLayout.PlayerColour.Red)
 End Select
 Next
 += 1
 += 1
 += 1
 += 1
 += 1
 += 1
 Purple) += 1
 += 1
 lblWwhiteCount.Text = count (CivLayout.PlayerColour.White)
 lblGreenCount.Text = count (CivLayout.PlayerColour.Green)
 lblBlueCount.Text
 lblYellowCount.Text
 = count (CivLayout.PlayerColour.Blue)
 = count (CivLayout.PlayerColour.Yellow)
 lblTurquoiseCount.Text
 = count (CivLayout.PlayerColour.Turquoise)
 lblOrangeCount.Text = count (CivLayout.PlayerColour.Orange)
 lblPurpleCount.Text
 = count (CivLayout.PlayerColour.Purple)
 lblRedCount.Text = count (CivLayout.PlayerColour.Red)
 End Sub
 Private
 Sub dgCities DataSourceChanged
 (ByVal sender As Object, ByVal e As System.
 EventArgs) Handles dgCities.DataSourceChanged
 If m drl Is Nothing Then ' UnitsOwned (City to Units)
 m
 UnitsOwned = False
 ElseIf
 Else
 m
 End If
 m
 CType (sender,
 DataGrid)
 m drl.RelationName.ToString
 m
 UnitsOwned = True
 UnitsOwned = False
 BuildFilter
 = False
 chkVet.Checked = False
 If
 embUnitType.Items.Count
 cmbUnitType.SelectedIndex
 End If
 rbAll.Checked = True
 BuildFilter
 m
 mVetFilter
 m UnitTypeFilter
 m UnitLocationFilter
 SetupUnitFilter
 End Sub
 Private
 = True
 = ""
 ()
 = ""
 Sub SetupUnitFilter()
 If m
 = ""
 UnitsOwned = True Then
 pnlUnitFilter.Visible
 Dim e As IEnumerator
 Dim i As Integer
 While e.MoveNext
 = 0
 .DataMember.ToString
 Then
 <> 0 Then
 = 0
 = True
 = CmbCity.Items.GetEnumerator
 =
 Tf e.Current Is m
 Civ.GetCity(dvCities.Item(dgCities.CurrentRowIndex).Item("
 CityNumber"))
 Then
 Exit While
 End
 If
 Tr o4= 1
 End While
 CmbCity.SelectedIndex
 = 1i
 CmbCity.Enabled = False
 Else
 pnlUnitFilter.Visible
 = False
 81
 ve
 ¢
 CmbCity.SelectedItem = "All"
 End
 End Sub
 Private
 If
 Sub FillUnitsPnlCount
 ()
 count.Clear (count, 0, count.Length)
 Dim r As DataRow
 For Each r In dsCiv.Tables
 ("Units")
 .Rows
 Select
 Case r.Item("OwnerColour®)
 Case "White"
 count (CivLayout.PlayerColour.White)
 Case "Green"
 count (CivLayout.PlayerColour.Green)
 Case "Blue"
 count (CivLayout.PlayerColour.Blue)
 Case "Yellow"
 count (CivLayout.PlayerColour.Yellow)
 Case "Turquoise"
 count (CivLayout.PlayerColour.Turquoise)
 Case "Orange"
 count (CivLayout.PlayerColour.Orange)
 Case "Purple"
 count (CivLayout.PlayerColour.
 Case "Red"
 count (CivLayout.PlayerColour.Red)
 End Select
 Next
 lblwhiteCount.Text
 += 1
 += 1
 += 1
 += 1
 += 1
 += 1
 Purple) += 1
 += 1
 = count (CivLayout.PlayerColour.White)
 lblGreenCount.Text
 = count (CivLayout.PlayerColour.Green)
 lblBlueCount.Text = count (CivLayout.PlayerColour.Blue)
 lblYellowCount.Text
 lblTurquoiseCount.Text
 = count (CivLayout.PlayerColour.Yellow)
 = count (CivLayout.PlayerColour.Turgquoise)
 lblOrangeCount.Text
 lblPurpleCount.Text
 = count (CivLayout.PlayerColour.Orange)
 = count (CivLayout.PlayerColour.Purple)
 lblRedCount.Text = count (CivLayout.PlayerColour.Red)
 End Sub
 #End Region
 #Region
 "Combo Box, Check Box and Radio Button Filters"
 Private
 Sub cmbOwnerColour SelectedIndexChanged
 (ByVal sender As System.Object,
 82
 ByVal e «¢
 As System.EventArgs) Handles cmbOwnerColour.SelectedIndexChanged
 mSupplyFilter
 mDemandFilter
 mProducingFilter
 mUnitTypeFilter
 m VetFilter
 mUnitCityFilter
 mUnitLocationFilter
 BuildCityFilter()
 BuildUnitFilter()
 BuildMapCellFilter()
 FillCityFilterCombos
 FillUnitTypeFilterCombo
 End Sub
 Private
 = ""
 = ""
 = ""
 = ""
 = "7
 = "7"
 = ""
 ()
 ()
 Sub cmbSupplies SelectedIndexChanged
 (ByVal sender As System.Object,
 System.EventArgs) Handles cmbSupplies.SelectedIndexChanged
 If m BuildFilter
 SupplyFilter
 RuildCityFilter
 End If
 End Sub
 Private
 = True Then
 ()
 {)
 Sub cmbDemands SelectedIndexChanged
 (ByVal sender As System.Object,
 System.EventArgs) Handles cmbDemands.SelectedIndexChanged
 If
 m
 BuildFilter
 DemandFilter
 BuildCityFilter
 End If
 = True Then
 ()
 ()
 ByVal e As «
 ByVal e As v4
 End Sub
 Private
 Sub cmbProducing SelectedIndexChanged
 (ByVal sender As System.Object,
 83
 ByVal e As «
 System.EventArgs) Handles cmbProducing.SelectedIndexChanged
 If
 End
 End Sub
 Private
 m
 m
 BuildFilter
 ProducingFilter
 BuildCityFilter()
 If
 = True Then
 ()
 Sub cmbUnitType SelectedIndexChanged
 BuildFilter
 (ByVal sender As System.Object,
 System.EventArgs) Handles cmbUnitType.SelectedIndexChanged
 If
 = True Then
 UnitTypeFilter()
 BuildUnitFilter()
 End If
 End Sub
 Private
 Sub chkVet CheckedChanged (ByVal sender As System.Object,
 EventArgs) Handles chkVet.CheckedChanged
 If mi
 End
 End Sub
 Private
 BuildFilter
 VetFilter()
 BuildUnitFilter
 If
 = True Then
 ()
 Sub CmbCity SelectedIndexChanged
 (ByVal sender As System.Object,
 System.EventArgs) Handles CmbCity.SelectedIndexChanged
 If
 m
 BuildFilter
 UnitCityFilter()
 BuildUnitFilter
 End If
 End Sub
 Private
 = True Then
 ()
 Sub rbAll CheckedChanged (ByVal sender As System.Object,
 EventArgs) Handles rbAll.CheckedChanged
 If
 rbAll.Checked
 UnitLocationFilter
 BuildUnitFilter
 End If
 End Sub
 Private
 = True And m
 ()
 ()
 Sub rbAway CheckedChanged
 EventArgs)
 BuildFilter
 Handles rbAway.CheckedChanged
 If
 rbAway.Checked
 UnitLocationFilter
 BuildUnitFilter()
 End If
 End Sub
 Private
 = True And m BuildFilter
 ()
 Sub rbHome CheckedChanged
 EventArgs)
 If
 Handles rbHome.CheckedChanged
 rbHome.Checked
 UnitLocationFilter{)
 BuildUnitFilter
 End If
 End Sub
 = True And m
 ()
 = True Then
 (ByVal sender As System.Object,
 = True Then
 BN
 (ByVal sender As System.Object,
 BuildFilter
 = True Then
 ByVal e As «¢
 ByVal e As System.
 ByVal e As
 ByVal e As System.
 ByVal e As System.
 ByVal e As System.
 Private Sub rbNone CheckedChanged (ByVal sender As System.Object, ByVal e As System.
 EventArgs) Handles rbNone.CheckedChanged
 If
 rbNone.Checked
 UnitLocationFilter()
 BuildUnitFilter()
 End If
 End Sub
 #End Region
 #Region "Filter
 = True And m
 and Helper Routines"
 BuildFilter
 = True Then
 ‘Shared Function CloneObject (ByVal ob] As Object) As Object
 '
 f
 ’
 Create
 a memory stream and formatter
 Dim mg As New MemoryStream (1000)
 Dim bf As New BinaryvFormatter
 "4
 v4
 ve
 ¥4
 v4
 v
 (Nothing,
 New StreamingContext
 '
 '
 (StreamingContextStates.Clone})
 Serialize the object into the stream
 '" bf.Serialize
 (ms, obj}
 ms.S%eek (0, SeekOrigin.Begin)
 '"'" Deserialzie
 ;
 :
 H
 CloneObject
 '
 Release
 ms.Close()
 "End Function
 Private
 into
 another
 = bf.Deserialize
 the
		Memory
 Sub BuildCityFilter
 object.
 ()
 OwnerFilter
 CityFilter
 m
 If
 m
 ()
 SupplyFilter
 = m
 OwnerFilter
 <> "" Then
 TE m-CikyFilter
 (ms)
 <> "7 Then
 mCityFilter
 End If
 m
 End If
 If
 &= " AND "
 CityFilter &= m
 SupplyFilter
 m
 DemandFilter
 If
 m
 CityFilter
 mCityFilter
 End If
 m CityFilter
 End If
 If
 m
 ProducingFilter
 If
 m
 CityFilter
 mCityFilter
 End If
 m
 <> "" Then
 <> "" Then
 &= " AND "
 &= m DemandFilter
 fe
 <> "" Then
 <> "" Then
 &= " AND "
 CityFilter &= m
 ProducingFilter
 End If
 If m
 CityPilter
 m CiltyFilter
 End If
 m CityFilter
 <> "" Then
 &= " BND *
 &= "CityNumber <> 255"
 dvCities.RowFilter = mCikyFilter
 lblCitiesCount.Text
 End Sub
 Private
 = dvCities.Count
 Sub BuildUnitFilter()
 OwnerFilter()
 m
 UnitFilter
 If
 m
 = m
 UnitTypeFilter
 OwnerFilter
 <> "" Then
 If m UnitFilter
 m UnitFilter
 End
 m
 End If
 If
 UnitFilter
 <> "" Then
 &= " AND *
 &= m
 UnitTypeFilter
 If m
 VeEFilter
 If
 m
 UnitFilter
 m UnitFilter
 End If
 m
 End If
 <> "* Then
 <> "" Then
 &= " AND "
 UnitFilter &= m
 VetFilter
 If
 m
 UnatCityFilter
 If m UnitFilter
 m
 End If
 UnitFilter
 <> "" Then
 <> "" Then
 &=" AND "
 m UnitFilter &= mUnitCityFilter
 End If
 If
 m
 UnitLocationFilter
 If
 Bnd If
 m
 UnitFilter
 End If
 iE
 m
 UnitFilter
 <> "" Then
 <> "" Then
 mUnitFilter
 &= " AND "
 &= m
 UnitLocationFilter
 dvUnits.RowFilter = m
 UnitFilter
 If
 m
 UnitsOwned
 Try
 = True Then
 mCM = CType (Me.BindingContext
 CurrencyManager)
 Tf m
 CM.Position
 =-1 Then
 (dgCities.DataSource,
 DRV.DataView.RowFilter = m
 UnitCityFilter
 84
 m
 dgCities.DataMember),
 mCM = CType (Me.BindingContext
 ,
 CurrencyManager)
 End If
 (dgCities.DataSource,
 m DRV = CType(m CM.Current, DataRowView)
 m DRV.DataView.RowFilter = IUnitFPiltex
 Catch ex As Exception
 MsgBox (ex.Message)
 End Try
 End If
 lbllUnitsCount.Text = dvUnits.Count
 End Sub
 Private
 If
 Sub BuildMapCellFilter
 ()
 cmbOwnerColour.SelectedIndex
 m MapCellFilter
 Else
 m MapCellFilter
 = ""
 = 0 Then
 dgCities.DataMember)
 = "MapCivColour = '" & cmbOwnerColour.SelectedItem.ToString
 14]
 End If
 dvMapCells.RowFilter = m
 MapCellFilter
 End Sub
 Private
 If
 Else
 End
 End Sub
 Private
 Sub OwnerFilter ()
 cmbOwnerColour.SelectedIndex
 m_ OwnerFilter
 m
 OwnerFilter
 = "7
 = 0 Then
 = "OwnerColour = '" &cmbOwnerColour.SelectedItem.ToString
 If
 Sub FillCityFilterCombos
 OwnerFilter()
 m
 CityFilter
 = m
 OwnerFilter
 ()
 dvCities.RowFilter = m
 CityFilter
 sc.Clear()
 dc.Clear()
 pr.Clear()
 Dim 1 As Integer
 For i = 0 To dvCities.Count- 1
 Dim r As DataRowView
 r = dvCities.Item(1)
 Dim s(2),
 d(2),
 p As String
 s(0) = r.Item("SuppliedCommodityl")
 s(l)
 = r.Item("SuppliedCommodity2")
 s{(2) = r.Item("SuppliedCommodity3")
 d(0) = r.Item("DemandedCommodityl")
 d(l)
 = r.Item("DemandedCommodityz2")
 d(2) = r.Item("DemandedCommodity3")
 Pp = r.Item("CityProducing")
 Dim J As Byte
 For
 7 = 0 To 2
 If Not (s(7]).StartsWith("("))
 If sc.BinarySearch(s(j))
 sc.Add(s(J))
 sc.Sort
 End
 End
 If
 If
 ()
 Tf Not (d(3).StartsWith("("))
 If dc.BinarySearch(d(j))
 dc.Add (d(j))
 dc.Sort
 End IT
 End If
 Next
 If pr.BinarySearch(p)
 pr.Add(p)
 pr.Sort()
 End
 Next
 If
 BuildFilter
 m
 cmbSupplies.BeginUpdate
 ()
 < 0 Then
 Then
 < 0 Then
 Then
 < 0 Then
 = False ' Stop filter from being built prematurely
 ()
 & "'"
 85
 ¢
 & "'\w
 cmbSupplies.Items.Clear()
 cmbSupplies.Items.Add ("All")
 Dim sca() As String
 ReDim sca(sc.Count
 sc.CopyTo (sca)- 1)
 cmbSupplies.Items.AddRange (sca)
 cmbSupplies.EndUpdate
 cmbSupplies.SelectedIndex
 cmbDemands.BeginUpdate
 cmbDemands.Items.Clear
 ()
 = 0
 ()
 ()
 cmbDemands.Items.Add ("All")
 Dim dca() As String
 ReDim dca{dc.Count
 dc.CopyTo (dca)- 1)
 cmbDemands . Items .AddRange (dca)
 cmbDemands . EndUpdate ()
 cmbDemands.SelectedIndex
 cmbProducing.BeginUpdate()
 cmbProducing.Items.Clear
 = 0
 ()
 cmbProducing.Items.Add ("All")
 Dim pra()
 As String
 ReDim pra(pr.Count
 pt .CopyTe (pra)
1)
 cmbProducing.Items.AddRange (pra)
 cmbProducing.EndUpdate
 cmbProducing.SelectedIndex
 rbAll.Checked = True
 m
 End Sub
 BuildFilter
 Private
 = True
 ()
 Sub FillUnitTypeFilterCombo
 OwnerFilter
 ()
 = 0
 m UnitFilter = mOwnerFilter
 dvUnits.RowFilter = m UnitFilter
 ut.Clear
 ()
 Dim 1 As Integer
 For i = 0 To dwUnits.Count- 1
 Dim r As DataRowView
 r
 = dvUnits.Item(i)
 Dim w BS SExring
 u = r.Item("UnitType"”)
 If
 ut.BinarySearch(u)
 ()
 < 0 Then
 ut.Add (u)
 ut.Sork
 End If
 Next
 ()
 cmbUnitType.BeginUpdate
 cmbUnitType.Items.Clear
 cmbUnitType.Items.
 Dim wed(} As String
 ReDim uta{ut.Count
 ut.CopyTo (uta)
 ()
 ()
 Add ("AL1L1l")
1)
 cmbUnitType.Items.AddRange (uta)
 cmbUnitType.EndUpdate
 nm
 BusldFilter
 ()
 = False * Turn Build filters off
 cmbUnitType.SelectedIndex
 i Bal ldFileer
 End Sub
 Private
 = True
 Sub SupplyFilter()
 If cmbSupplies.SelectedIndex
 supplyFilter
 Wm
 Else
 m
 SupplyFilter
 = 0
 = ""
 = 0 Then
 = "(SuppliedCommodityl = '" &
 cmbSupplies.SelectedItem.ToString
 & "! OR SuppliedCommodity?2 = mE
 cmbSupplies.SelectedItem.ToString
 & "' ORSuppliedCommodity3 = '" &
 cmbSupplies.SelectedItem.ToString & "')"
 BiG: Ef
 End Sub
 Sub DemandFilter()
 86
 Private
 If
 Else
 cmbDemands.SelectedIndex
 m DemandFilter
 mDemandFilter
 = Ue
 = 0 Then
 = "(DemandedCommodityl = '" &
 cmbDemands.SelectedItem.ToString
 & "' OR DemandedCommodity2 = '" &
 cmbDemands.SelectedItem.ToString
 cmbDemands.SelectedItem.ToString
 End If
 End Sub
 Private
 If
 Else
 Sub ProducingFilter()
 cmbProducing.SelectedIndex
 mProducingFilter
 m
 ProducingFilter
 = 0 Then
 = "V
 & "' ORDemandedCommodity3 = '" &
 & "")"
 = "CityProducing = '" &
 cmbProducing.SelectedItem.ToString
 End If
 End Sub
 Private
 Sub UnitTypeFilter()
 If cmbUnitType.SelectedIndex
 wm Ung Clyper:
 Else
 mr
 UniETypeFilter
 = 0 Then
 liter
 = ""
 & "'
 = "UnitType = '" & cmbUnitlype.delecteditem.TeoString
 Bnd IT
 End Sub
 Private
 If
 Else
 Sub VetFilter()
 clik¥Yek.Chechked
 m VeEtRilter
 mh
 YekPllter
 Bnd IE
 End Sub
 Private
 False
 = "7
 = "Veferan
 Then
 = "Trus'"™
 Sub UnitCityFilter()
 If TypeOf (CmbCity.SelectedIlItem)
 wm
 Else
 PniECityFilter
 = "7
 Is String Then ' All
 Dim ¢ As City = CType(CmbCity.SelectedItem(),
 Mm
 UnitCityFilter
 End If
 End Sub
 Private
 = "HomeCityNumber = * & c.CityNumber.ToString
 Sub UnitLocationFilter
 ()
 If rbHome.Checked = True Then
 m
 ElseIf
 nm
 ElseIf
 UnitLocationFilter
 = "UnitLocation = 'Home'"
 rbAway.Checked = True Then
 UnithecatictiFilter
 rbNone.Checked
 m UnitLocationFilter
 —
 Else ' All
 th UnithecatienPilter
 End
 Enea Sub
 #End Region
 End Class
 IT
 = "UnitLecatien = 'Away'"
 = True Then
 = "UnitLocation
 = "V
 = 'None'"”
 & "'"
 City)
 87