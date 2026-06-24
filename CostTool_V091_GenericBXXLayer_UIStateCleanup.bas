Option Compare Database
Option Explicit

' ============================================================
' Decommissioning Cost Tool - Access Builder v0.9.1 Add-On
' Generic BXX Template Layer + UI Input Polish
' ============================================================
'
' Purpose:
'   Adds a universal high-level estimating layer on top of the stable v0.6.2
'   bottom-up Access cost tool.
'
' Core idea:
'   New building inputs -> generic BXX-style WBS estimate -> generated lines
'   -> existing tblJobLines fine-tuning screen.
'
' Design principles:
'   - Requires stable v0.6.2 baseline objects.
'   - Does NOT overwrite tblCostLibrary master records.
'   - Uses table-driven templates instead of hard-coded B56/B18 logic.
'   - Supports a generic Building / Structure D&D basis and a generic
'     Soil / Outdoor Remediation basis.
'   - Exposes key universal inputs: area, duration, crew size, class, scope basis,
'     removal adjustment, percent clean/contaminated, removal depth, backfill depth,
'     and the judgement levers already present in the Excel template.
'
' v0.9.1 UI / input polish:
'   - No calculation changes to the validated v0.8.9 engine.
'   - New input rows are created blank instead of pre-populated with B56-like values.
'   - Management staff number inputs are exposed alongside use-factor inputs.
'   - Site preparation can be selected by individual Excel task checkbox while
'     preserving the original 16 hr x 4 specialists + 16 hr x 1 PM calculation.
'
' Entry point:
'   BuildCostTool_V091_GenericBXXLayer
'
' ============================================================

Public Sub BuildCostTool_V091_GenericBXXLayer()
    On Error GoTo ErrHandler

    DoCmd.Hourglass True
    Application.Echo False

    V091_CheckBaseObjects
    V091_CreateTables
    V091_UpdateBaseSchema
    V091_EnsureModelAllowanceLibraryItems
    V091_SeedSettings
    V091_SeedEquipmentTemplates
    V091_SeedConsumableTemplates
    V091_SeedAreaActivityTemplates
    V091_SeedDefaultInputs
    V091_CreateQueries

    V091_DeleteFormIfExists "frmV091GeneratedLinesSubform"
    V091_DeleteFormIfExists "frmV091GenericEstimate"
    V091_DeleteFormIfExists "frmV091PortfolioOverview"

    V091_CreateGeneratedLinesSubform
    V091_CreateEstimateForm
    V091_CreatePortfolioForm

    Application.Echo True
    DoCmd.Hourglass False

    DoCmd.OpenForm "frmV091GenericEstimate"

    MsgBox "Cost Tool v0.9.1 UI-state-polished generic BXX layer built successfully." & vbCrLf & vbCrLf & _
           "Use frmV091GenericEstimate as the universal new-building entry point.", _
           vbInformation, "v0.9.1 Build Complete"
    Exit Sub

ErrHandler:
    Application.Echo True
    DoCmd.Hourglass False
    MsgBox "v0.9.1 build failed: " & Err.Number & vbCrLf & Err.Description, _
           vbCritical, "v0.9.1 Build Error"
End Sub

Private Sub V091_CheckBaseObjects()
    If Not V091_TableExists("tblJobs") Then Err.Raise vbObjectError + 8201, , "tblJobs not found. Run BuildCostTool_V062_AllInOne first."
    If Not V091_TableExists("tblJobLines") Then Err.Raise vbObjectError + 8202, , "tblJobLines not found. Run BuildCostTool_V062_AllInOne first."
    If Not V091_TableExists("tblCostLibrary") Then Err.Raise vbObjectError + 8203, , "tblCostLibrary not found. Run BuildCostTool_V062_AllInOne first."
    If Not V091_TableExists("tblCategories") Then Err.Raise vbObjectError + 8204, , "tblCategories not found. Run BuildCostTool_V062_AllInOne first."
    If Not V091_QueryExists("qryGrandTotals") Then Err.Raise vbObjectError + 8205, , "qryGrandTotals not found. Run BuildCostTool_V062_AllInOne first."
End Sub

' ============================================================
' TABLES
' ============================================================

Private Sub V091_CreateTables()
    Dim sql As String

    sql = "CREATE TABLE tblV091BuildingInputs ("
    sql = sql & "JobID TEXT(50) CONSTRAINT pk_tblV091BuildingInputs PRIMARY KEY, "
    sql = sql & "BuildingCode TEXT(50), "
    sql = sql & "BuildingName TEXT(255), "
    sql = sql & "FacilityClass TEXT(10), "
    sql = sql & "FacilityType TEXT(100), "
    sql = sql & "EstimateBasis TEXT(100), "
    sql = sql & "TotalAreaM2 DOUBLE, "
    sql = sql & "FootprintAreaM2 DOUBLE, "
    sql = sql & "ProjectDurationDays DOUBLE, "
    sql = sql & "WorkDays DOUBLE, "
    sql = sql & "CrewSize DOUBLE, "
    sql = sql & "ScaleRemovalByCrew YESNO, "
    sql = sql & "RemovalAdjustmentPct DOUBLE, "
    sql = sql & "PortfolioManagerUseFactor DOUBLE, PortfolioManagerNumber DOUBLE, "
    sql = sql & "SeniorPMUseFactor DOUBLE, SeniorPMNumber DOUBLE, "
    sql = sql & "ProjectManagerUseFactor DOUBLE, ProjectManagerNumber DOUBLE, "
    sql = sql & "ProcedureHours DOUBLE, "
    sql = sql & "QASafetyHours DOUBLE, "
    sql = sql & "IncludeTraining YESNO, "
    sql = sql & "IncludeSitePrep YESNO, "
    sql = sql & "SitePrepTaskCount DOUBLE, SitePrepHoursPerTask DOUBLE, SitePrepInitialSurvey YESNO, SitePrepBoundariesHepa YESNO, SitePrepStagingArea YESNO, SitePrepRadSegregation YESNO, SitePrepElectricalIsolation YESNO, SitePrepPipingIsolation YESNO, "
    sql = sql & "IncludeDetailedCharacterization YESNO, "
    sql = sql & "CharacterizationSpecialistCount DOUBLE, "
    sql = sql & "CharacterizationPMCount DOUBLE, "
    sql = sql & "CharacterizationHoursPerPerson DOUBLE, "
    sql = sql & "IncludeConsumables YESNO, "
    sql = sql & "PercentClean DOUBLE, "
    sql = sql & "PercentContaminated DOUBLE, "
    sql = sql & "RemovalDepthM DOUBLE, "
    sql = sql & "BackfillDepthM DOUBLE, "
    sql = sql & "AsbestosPipeLengthM DOUBLE, "
    sql = sql & "AsbestosTileAreaM2 DOUBLE, "
    sql = sql & "ConsumableMonths DOUBLE, "
    sql = sql & "DosimeterYears DOUBLE, "
    sql = sql & "BioassayYears DOUBLE, "
    sql = sql & "IsRadiological YESNO, "
    sql = sql & "IncludeHotCell YESNO, "
    sql = sql & "HotCellAreaM2 DOUBLE, "
    sql = sql & "Notes MEMO, "
    sql = sql & "LastGeneratedAt DATETIME, "
    sql = sql & "UpdatedAt DATETIME"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091BuildingInputs", sql

    sql = "CREATE TABLE tblV091Settings ("
    sql = sql & "SettingName TEXT(100) CONSTRAINT pk_tblV091Settings PRIMARY KEY, "
    sql = sql & "SettingValueNumber DOUBLE, "
    sql = sql & "SettingValueText TEXT(255), "
    sql = sql & "Notes MEMO, "
    sql = sql & "UpdatedAt DATETIME"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091Settings", sql

    sql = "CREATE TABLE tblV091EquipmentTemplate ("
    sql = sql & "TemplateID AUTOINCREMENT CONSTRAINT pk_tblV091EquipmentTemplate PRIMARY KEY, "
    sql = sql & "FacilityClass TEXT(10), "
    sql = sql & "CategoryName TEXT(100), "
    sql = sql & "ItemID LONG, "
    sql = sql & "ItemName TEXT(255), "
    sql = sql & "Quantity DOUBLE, "
    sql = sql & "DefaultEnabled YESNO, "
    sql = sql & "Notes MEMO"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091EquipmentTemplate", sql

    sql = "CREATE TABLE tblV091ConsumableTemplate ("
    sql = sql & "TemplateID AUTOINCREMENT CONSTRAINT pk_tblV091ConsumableTemplate PRIMARY KEY, "
    sql = sql & "FacilityClass TEXT(10), "
    sql = sql & "CategoryName TEXT(100), "
    sql = sql & "ItemID LONG, "
    sql = sql & "ItemName TEXT(255), "
    sql = sql & "UseRate DOUBLE, "
    sql = sql & "DurationBasis TEXT(50), "
    sql = sql & "DefaultEnabled YESNO, "
    sql = sql & "Notes MEMO"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091ConsumableTemplate", sql

    sql = "CREATE TABLE tblV091StaffUseTemplate ("
    sql = sql & "TemplateID AUTOINCREMENT CONSTRAINT pk_tblV091StaffUseTemplate PRIMARY KEY, "
    sql = sql & "FacilityClass TEXT(10), "
    sql = sql & "RoleName TEXT(255), "
    sql = sql & "LabourItemID LONG, "
    sql = sql & "HoursPerDay DOUBLE, "
    sql = sql & "UseFactor DOUBLE, "
    sql = sql & "DefaultEnabled YESNO, "
    sql = sql & "Notes MEMO"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091StaffUseTemplate", sql

    sql = "CREATE TABLE tblV091AreaActivityTemplate ("
    sql = sql & "TemplateID AUTOINCREMENT CONSTRAINT pk_tblV091AreaActivityTemplate PRIMARY KEY, "
    sql = sql & "EstimateBasis TEXT(100), "
    sql = sql & "WBSCode TEXT(50), "
    sql = sql & "WBSSubCode TEXT(50), "
    sql = sql & "Description TEXT(255), "
    sql = sql & "CategoryName TEXT(100), "
    sql = sql & "ItemID LONG, "
    sql = sql & "QuantitySource TEXT(100), "
    sql = sql & "UnitName TEXT(50), "
    sql = sql & "UnitRateAUD DOUBLE, "
    sql = sql & "ApplyRemovalAdjustment YESNO, "
    sql = sql & "RequiresRadiological YESNO, "
    sql = sql & "CleanVolRateM3 DOUBLE, "
    sql = sql & "LsaVolRateM3 DOUBLE, "
    sql = sql & "HazVolRateM3 DOUBLE, "
    sql = sql & "MixedVolRateM3 DOUBLE, "
    sql = sql & "DefaultEnabled YESNO, "
    sql = sql & "Notes MEMO"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091AreaActivityTemplate", sql

    sql = "CREATE TABLE tblV091GeneratedLines ("
    sql = sql & "GeneratedLineID AUTOINCREMENT CONSTRAINT pk_tblV091GeneratedLines PRIMARY KEY, "
    sql = sql & "JobID TEXT(50) NOT NULL, "
    sql = sql & "WBSCode TEXT(50), "
    sql = sql & "WBSSubCode TEXT(50), "
    sql = sql & "Description TEXT(255), "
    sql = sql & "LineType TEXT(50), "
    sql = sql & "TargetCategoryName TEXT(100), "
    sql = sql & "TargetItemID LONG, "
    sql = sql & "QuantityPhysical DOUBLE, "
    sql = sql & "UnitName TEXT(50), "
    sql = sql & "UnitRateAUD CURRENCY, "
    sql = sql & "AdjustmentFactor DOUBLE, "
    sql = sql & "GeneratedCostAUD CURRENCY, "
    sql = sql & "CleanVolM3 DOUBLE, "
    sql = sql & "LsaVolM3 DOUBLE, "
    sql = sql & "HazVolM3 DOUBLE, "
    sql = sql & "MixedVolM3 DOUBLE, "
    sql = sql & "SourceNote MEMO, "
    sql = sql & "AppliedToJobLine YESNO, "
    sql = sql & "CreatedAt DATETIME"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091GeneratedLines", sql

    sql = "CREATE TABLE tblV091JobLineBackup ("
    sql = sql & "BackupID AUTOINCREMENT CONSTRAINT pk_tblV091JobLineBackup PRIMARY KEY, "
    sql = sql & "BackupAt DATETIME, "
    sql = sql & "JobID TEXT(50), "
    sql = sql & "JobLineID LONG, "
    sql = sql & "CategoryName TEXT(100), "
    sql = sql & "ItemID LONG, "
    sql = sql & "IncludeItem YESNO, "
    sql = sql & "Quantity DOUBLE, "
    sql = sql & "BaseUnitRateUSD2009 CURRENCY"
    sql = sql & ");"
    V091_CreateTableIfMissing "tblV091JobLineBackup", sql
End Sub

Private Sub V091_UpdateBaseSchema()
    ' v0.9.1 can be re-run over an earlier v0.9.1 build without losing data.
    If V091_TableExists("tblV091BuildingInputs") Then
        V091_AddFieldIfMissing "tblV091BuildingInputs", "PortfolioManagerNumber", "DOUBLE"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SeniorPMNumber", "DOUBLE"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "ProjectManagerNumber", "DOUBLE"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepHoursPerTask", "DOUBLE"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepInitialSurvey", "YESNO"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepBoundariesHepa", "YESNO"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepStagingArea", "YESNO"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepRadSegregation", "YESNO"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepElectricalIsolation", "YESNO"
        V091_AddFieldIfMissing "tblV091BuildingInputs", "SitePrepPipingIsolation", "YESNO"
    End If

    V091_AddFieldIfMissing "tblJobLines", "V091GeneratedQuantity", "DOUBLE"
    V091_AddFieldIfMissing "tblJobLines", "V091GeneratedCostAUD", "CURRENCY"
    V091_AddFieldIfMissing "tblJobLines", "V091IsGenerated", "YESNO"
    V091_AddFieldIfMissing "tblJobLines", "V091OriginalBaseRateUSD2009", "CURRENCY"
    V091_AddFieldIfMissing "tblJobLines", "V091LastGeneratedAt", "DATETIME"
End Sub

' ============================================================
' SEED DATA
' ============================================================

Private Sub V091_SeedSettings()
    If DCount("*", "tblV091Settings") > 0 Then Exit Sub

    V091_AddSettingNum "EquipmentEscalation2009To2026", 1.58014, "Generic template: 2009 to 2026 escalation factor for equipment/materials before AUD conversion."
    V091_AddSettingNum "UsdToAudExchangeRate", 1.54, "Generic template: Inputs!C37 exchange factor applied to the whole 1.1.2 equipment/materials/consumables block when the workbook basis is US."
    V091_AddSettingNum "SmallToolsPctOfActivityLabor", 0.02, "Generic template: Small Tools = 2% of activity labour/activity cost base."
    V091_AddSettingNum "HpEquipmentReplacementPctOfActivityLabor", 0.05, "Generic template: HP Equipment Replacement = 5%."
    V091_AddSettingNum "EquipmentOverheadPct", 0.08, "Generic template: DGC OH&P on equipment/materials = 8%."
    V091_AddSettingNum "WasteDensityLbPerCuft", 100, "Generic template waste density."
    V091_AddSettingNum "M3ToCuft", 35.3147, "Generic template conversion."
    V091_AddSettingNum "LbPerWasteTon", 2000, "Short ton conversion used by template."
    V091_AddSettingNum "IndustrialWasteRate", 300, "Clean / industrial waste disposal rate."
    V091_AddSettingNum "HazardousWasteRate", 500, "Hazardous and mixed waste disposal rate."
End Sub

Private Sub V091_AddSettingNum(ByVal settingName As String, ByVal settingValue As Double, ByVal notes As String)
    CurrentDb.Execute "INSERT INTO tblV091Settings (SettingName, SettingValueNumber, Notes, UpdatedAt) VALUES (" & _
        V091_Q(settingName) & ", " & V091_SqlNum(settingValue) & ", " & V091_Q(notes) & ", Now());", dbFailOnError
End Sub

Private Sub V091_SeedEquipmentTemplates()
    If DCount("*", "tblV091EquipmentTemplate") > 0 Then Exit Sub

    ' B1 generic BXX class equipment template.
    V091_AddEquip "B1", 1, "HEPA filter systems", 2
    V091_AddEquip "B1", 2, "Replacement filters", 24
    V091_AddEquip "B1", 3, "Respirator", 12
    V091_AddEquip "B1", 4, "Rad/Vac wet-dry high eff. vacuum", 1
    V091_AddEquip "B1", 5, "Rad/Vac wet-dry high eff. vacuum filters", 5
    V091_AddEquip "B1", 6, "Reciprocating Saws", 5
    V091_AddEquip "B1", 7, "Pneumatic chipping hammers", 2
    V091_AddEquip "B1", 8, "Chipping hammer blades", 20
    V091_AddEquip "B1", 9, "Purchase an air compressor", 2
    V091_AddEquip "B1", 10, "Jackhammer", 4
    V091_AddEquip "B1", 11, "Jackhammer Chisels", 10
    V091_AddEquip "B1", 12, "Safety glasses", 15
    V091_AddEquip "B1", 13, "Fall protection - harness", 6
    V091_AddEquip "B1", 14, "Fall protection - lanyard", 6
    V091_AddEquip "B1", 15, "Hardhats", 15
    V091_AddEquip "B1", 16, "Hard hat hearing protection", 15
    V091_AddEquip "B1", 17, "Trailer rental", 1
    V091_AddEquip "B1", 18, "Phone and computer hook-up", 1
    V091_AddEquip "B1", 19, "Industrial hygiene instrumentation", 0
    V091_AddEquip "B1", 20, "Tractor loader, wheeled", 1
    V091_AddEquip "B1", 21, "Wheeled skid steer with concrete hammer", 1
    V091_AddEquip "B1", 22, "Excavator", 1
    V091_AddEquip "B1", 23, "Floor shaver", 1
    V091_AddEquip "B1", 24, "Wall shaver", 1
    V091_AddEquip "B1", 25, "Man lift", 2
    V091_AddEquip "B1", 26, "Crane rental for building removal", 1

    ' C1 and D1 use the same editable template as generic radiological classes.
    V091_CopyEquipmentClass "B1", "C1"
    V091_CopyEquipmentClass "B1", "D1"
    CurrentDb.Execute "UPDATE tblV091EquipmentTemplate SET Quantity=2 WHERE FacilityClass='C1' AND ItemID=4;", dbFailOnError

    ' D2 non-radiological generic equipment template.
    V091_AddEquip "D2", 1, "HEPA filter systems", 0
    V091_AddEquip "D2", 2, "Replacement filters", 0
    V091_AddEquip "D2", 3, "Respirator", 0
    V091_AddEquip "D2", 4, "Rad/Vac wet-dry high eff. vacuum", 0
    V091_AddEquip "D2", 5, "Rad/Vac wet-dry high eff. vacuum filters", 0
    V091_AddEquip "D2", 6, "Reciprocating Saws", 0
    V091_AddEquip "D2", 7, "Pneumatic chipping hammers", 2
    V091_AddEquip "D2", 8, "Chipping hammer blades", 20
    V091_AddEquip "D2", 9, "Purchase an air compressor", 2
    V091_AddEquip "D2", 10, "Jackhammer", 4
    V091_AddEquip "D2", 11, "Jackhammer Chisels", 10
    V091_AddEquip "D2", 12, "Safety glasses", 15
    V091_AddEquip "D2", 13, "Fall protection - harness", 0
    V091_AddEquip "D2", 14, "Fall protection - lanyard", 0
    V091_AddEquip "D2", 15, "Hardhats", 15
    V091_AddEquip "D2", 16, "Hard hat hearing protection", 15
    V091_AddEquip "D2", 17, "Trailer rental", 1
    V091_AddEquip "D2", 18, "Phone and computer hook-up", 1
    V091_AddEquip "D2", 19, "Industrial hygiene instrumentation", 0
    V091_AddEquip "D2", 20, "Tractor loader, wheeled", 1
    V091_AddEquip "D2", 21, "Wheeled skid steer with concrete hammer", 1
    V091_AddEquip "D2", 22, "Excavator", 1
    V091_AddEquip "D2", 23, "Floor shaver", 0
    V091_AddEquip "D2", 24, "Wall shaver", 0
    V091_AddEquip "D2", 25, "Man lift", 0
    V091_AddEquip "D2", 26, "Crane rental for building removal", 0
End Sub

Private Sub V091_AddEquip(ByVal facilityClass As String, ByVal itemID As Long, ByVal itemName As String, ByVal qty As Double)
    CurrentDb.Execute "INSERT INTO tblV091EquipmentTemplate (FacilityClass, CategoryName, ItemID, ItemName, Quantity, DefaultEnabled, Notes) VALUES (" & _
        V091_Q(facilityClass) & ", 'Equipment', " & itemID & ", " & V091_Q(itemName) & ", " & V091_SqlNum(qty) & ", True, 'Seeded from generic BXX class equipment template.');", dbFailOnError
End Sub

Private Sub V091_CopyEquipmentClass(ByVal sourceClass As String, ByVal targetClass As String)
    CurrentDb.Execute "INSERT INTO tblV091EquipmentTemplate (FacilityClass, CategoryName, ItemID, ItemName, Quantity, DefaultEnabled, Notes) " & _
        "SELECT " & V091_Q(targetClass) & ", CategoryName, ItemID, ItemName, Quantity, DefaultEnabled, 'Copied generic class template from ' & FacilityClass & '; edit table if required.' " & _
        "FROM tblV091EquipmentTemplate WHERE FacilityClass=" & V091_Q(sourceClass) & ";", dbFailOnError
End Sub

Private Sub V091_SeedConsumableTemplates()
    If DCount("*", "tblV091ConsumableTemplate") > 0 Then Exit Sub

    V091_AddConsumableAllRad 27, "Coveralls", 2, "man-day"
    V091_AddConsumableAllRad 28, "Hoods", 0.1, "man-day"
    V091_AddConsumableAllRad 29, "Shoe covers", 4, "man-day"
    V091_AddConsumableAllRad 30, "Latex gloves", 4, "man-day"
    V091_AddConsumableAllRad 31, "Rubber overshoes", 0.01, "man-day"
    V091_AddConsumableAllRad 32, "Gloves", 0.01, "man-day"
    V091_AddConsumableAllRad 33, "Dosimeters", 1, "man-year"
    V091_AddConsumableAllRad 34, "TLDs", 1, "man-month"
    V091_AddConsumableAllRad 35, "Bioassays", 2, "man-year-bioassay"

    V091_AddConsumableClass "D2", 27, "Coveralls", 0, "man-day"
    V091_AddConsumableClass "D2", 28, "Hoods", 0, "man-day"
    V091_AddConsumableClass "D2", 29, "Shoe covers", 0, "man-day"
    V091_AddConsumableClass "D2", 30, "Latex gloves", 0, "man-day"
    V091_AddConsumableClass "D2", 31, "Rubber overshoes", 0, "man-day"
    V091_AddConsumableClass "D2", 32, "Gloves", 0, "man-day"
    V091_AddConsumableClass "D2", 33, "Dosimeters", 0, "man-year"
    V091_AddConsumableClass "D2", 34, "TLDs", 0, "man-month"
    V091_AddConsumableClass "D2", 35, "Bioassays", 0, "man-year-bioassay"
End Sub

Private Sub V091_AddConsumableAllRad(ByVal itemID As Long, ByVal itemName As String, ByVal useRate As Double, ByVal basis As String)
    V091_AddConsumableClass "B1", itemID, itemName, useRate, basis
    V091_AddConsumableClass "C1", itemID, itemName, useRate, basis
    V091_AddConsumableClass "D1", itemID, itemName, useRate, basis
End Sub

Private Sub V091_AddConsumableClass(ByVal facilityClass As String, ByVal itemID As Long, ByVal itemName As String, ByVal useRate As Double, ByVal basis As String)
    CurrentDb.Execute "INSERT INTO tblV091ConsumableTemplate (FacilityClass, CategoryName, ItemID, ItemName, UseRate, DurationBasis, DefaultEnabled, Notes) VALUES (" & _
        V091_Q(facilityClass) & ", 'Consumables', " & itemID & ", " & V091_Q(itemName) & ", " & V091_SqlNum(useRate) & ", " & V091_Q(basis) & ", True, 'Seeded from generic BXX consumable template.');", dbFailOnError
End Sub

Private Sub V091_SeedStaffUseTemplates()
    If DCount("*", "tblV091StaffUseTemplate") > 0 Then Exit Sub
    ' Not called directly in build; kept available for future table editing.
End Sub

Private Sub V091_SeedAreaActivityTemplates()
    If DCount("*", "tblV091AreaActivityTemplate") > 0 Then Exit Sub

    ' Generic Building / Structure D&D route. UnitRateAUD values are Summary 2026 base rates before crew-factor scaling.
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.1", "Remove contaminated asbestos pipe insulation", "Area Costs", 1, "AsbestosPipeLengthM", "m", 721.8449512, False, False, 0, 0, 0, 0.04615349, "Generic BXX asbestos pipe activity. Excel parity v0.9.1: contaminated pipe insulation waste routes to mixed waste, not clean waste."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.1", "Remove asbestos tile", "Area Costs", 2, "AsbestosTileAreaM2", "m2", 58.25179202, False, False, 0.03061509, 0, 0, 0, "Generic BXX asbestos tile activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.2", "System Clean", "Area Costs", 3, "TotalAreaM2", "m2", 603.9881484, True, False, 0.210971851, 0, 0, 0, "Generic BXX system clean activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.2", "System LLW", "Area Costs", 4, "TotalAreaM2", "m2", 414.5889597, True, True, 0, 0.021826187, 0, 0, "Generic BXX radiological system LLW activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.2", "System Hazardous", "Area Costs", 5, "TotalAreaM2", "m2", 11.19679335, True, False, 0, 0, 0.000347973, 0, "Generic BXX system hazardous activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.2", "System Mixed", "Area Costs", 6, "TotalAreaM2", "m2", 8.8206536, True, True, 0, 0, 0, 0.002881651, "Generic BXX radiological mixed-system activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.3", "Decon Cleaning", "Area Costs", 7, "TotalAreaM2", "m2", 18.15004985, True, False, 0.009905621, 0, 0, 0, "Generic BXX decon clean activity."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.3", "Decontaminate Hot Cells", "Area Costs", 31, "HotCellAreaM2", "m2", 120.8368922, True, True, 0, 0.002542516119, 0, 0, "Generic hot-cell decontamination activity. Quantity defaults to total area when hot-cell scope is included and no hot-cell area is entered."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.3", "Remove hazardous material", "Model Allowances", 7, "TotalAreaM2", "m2", 21.15184624, True, False, 0, 0, 0.2726278792, 0, "Waste-volume carrier for hazardous material. Cost is activated only when hot-cell scope is included; otherwise rate is set to zero to match non-hot-cell Excel tabs."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.3", "Decon Contaminated", "Area Costs", 9, "TotalAreaM2", "m2", 427.7725506, True, True, 0, 0.000844498, 0, 0, "Generic BXX radiological contaminated decon."
    V091_AddAreaActivity "Building / Structure D&D", "1.5", "1.5.4", "Final Survey", "Area Costs", 34, "TotalAreaM2", "m2", 242.3439213, False, False, 0, 0, 0, 0, "Generic BXX final survey."
    V091_AddAreaActivity "Building / Structure D&D", "1.6", "1.6.1", "Remove Building", "Area Costs", 35, "TotalAreaM2", "m2", 455.2381391, True, False, 0.770874078, 0, 0, 0, "Generic BXX remove building."
    V091_AddAreaActivity "Building / Structure D&D", "1.6", "1.6.2", "Remove Hot Cell", "Area Costs", 31, "HotCellAreaM2", "m2", 150.7114894, True, True, 0.1222098525, 0, 0, 0, "Generic hot-cell removal activity. Quantity defaults to total area when hot-cell scope is included and no hot-cell area is entered."
    V091_AddAreaActivity "Building / Structure D&D", "1.6", "1.6.3", "Grade and Seed", "Area Costs", 36, "GradeAreaM2", "m2", 15.82447093, False, False, 0, 0, 0, 0, "Generic BXX grade and seed at 2x footprint."
    V091_AddAreaActivity "Building / Structure D&D", "1.6", "1.6.4", "Backfill", "Area Costs", 37, "BackfillM3", "m3", 23.78663385, False, False, 0, 0, 0, 0, "Generic BXX backfill using SI depth."

    ' Generic Soil / Outdoor Remediation route. UnitRateAUD values are Summary 2026 base rates before crew-factor scaling.
    ' v0.9.1 Excel parity cleanup:
    '   B18-style soil/outdoor remediation is generated as one direct WBS 1.5.1
    '   Remove Soil and Asphalt activity, rather than approximating it through
    '   legacy System Hazardous + System Mixed + Soil rows.
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.5", "1.5.1", "Remove Soil and Asphalt", "Area Costs", 28, "TotalAreaM2", "m2", 85.25913043, True, False, 0, 0, 0, 0, "Excel parity v0.9.1 soil/outdoor route: direct WBS 1.5.1 Remove Soil and Asphalt line."
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.5", "1.5.1", "Clean soil/asphalt waste volume carrier", "Model Allowances", 7, "CleanSoilVolumeM3", "m3", 0, False, False, 1, 0, 0, 0, "Excel parity v0.9.1 soil/outdoor route: zero-cost clean soil/asphalt volume carrier for waste disposal."
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.5", "1.5.1", "Contaminated soil/asphalt tracked volume", "Model Allowances", 7, "ContaminatedSoilVolumeM3", "m3", 0, False, False, 0, 1, 0, 0, "Excel parity v0.9.1 soil/outdoor route: zero-cost contaminated soil/asphalt volume tracker; cost rate remains zero in template."
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.5", "1.5.4", "Final Survey", "Area Costs", 34, "TotalAreaM2", "m2", 242.3439213, False, False, 0, 0, 0, 0, "Generic soil/outdoor final survey rate."
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.6", "1.6.3", "Grade and Seed", "Area Costs", 36, "TotalAreaM2", "m2", 15.82447093, False, False, 0, 0, 0, 0, "Generic soil/outdoor grade and seed."
    V091_AddAreaActivity "Soil / Outdoor Remediation", "1.6", "1.6.4", "Backfill", "Area Costs", 37, "BackfillM3", "m3", 23.78663385, False, False, 0, 0, 0, 0, "Generic soil/outdoor backfill using SI depth."
End Sub

Private Sub V091_AddAreaActivity(ByVal basis As String, ByVal wbsCode As String, ByVal wbsSub As String, ByVal descText As String, ByVal cat As String, ByVal itemID As Long, ByVal source As String, ByVal unitName As String, ByVal rate As Double, ByVal applyAdj As Boolean, ByVal reqRad As Boolean, ByVal cleanRate As Double, ByVal lsaRate As Double, ByVal hazRate As Double, ByVal mixedRate As Double, ByVal notes As String)
    CurrentDb.Execute "INSERT INTO tblV091AreaActivityTemplate (EstimateBasis, WBSCode, WBSSubCode, Description, CategoryName, ItemID, QuantitySource, UnitName, UnitRateAUD, ApplyRemovalAdjustment, RequiresRadiological, CleanVolRateM3, LsaVolRateM3, HazVolRateM3, MixedVolRateM3, DefaultEnabled, Notes) VALUES (" & _
        V091_Q(basis) & ", " & V091_Q(wbsCode) & ", " & V091_Q(wbsSub) & ", " & V091_Q(descText) & ", " & V091_Q(cat) & ", " & itemID & ", " & V091_Q(source) & ", " & V091_Q(unitName) & ", " & V091_SqlNum(rate) & ", " & IIf(applyAdj, "True", "False") & ", " & IIf(reqRad, "True", "False") & ", " & V091_SqlNum(cleanRate) & ", " & V091_SqlNum(lsaRate) & ", " & V091_SqlNum(hazRate) & ", " & V091_SqlNum(mixedRate) & ", True, " & V091_Q(notes) & ");", dbFailOnError
End Sub

Private Sub V091_SeedDefaultInputs()
    ' v0.9.1 UI/state cleanup:
    ' Do NOT create input shells for existing tblJobs during build.
    ' This prevents the form from opening on JOB-001 or any other old job.
    ' Input rows are created only when the user clicks New / Attach Job or
    ' explicitly selects an existing job from the Load Job combo.
End Sub

' ============================================================
' MODEL ALLOWANCE CATEGORY / LIBRARY ITEMS
' ============================================================

Private Sub V091_EnsureModelAllowanceLibraryItems()
    If DCount("*", "tblCategories", "CategoryName='Model Allowances'") = 0 Then
        CurrentDb.Execute "INSERT INTO tblCategories (CategoryName, DisplayOrder) VALUES ('Model Allowances', 6);", dbFailOnError
    End If

    V091_EnsureLibraryItem "Model Allowances", 1, "1.1", "1.1.2", "Small Tools Allowance", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 2, "1.1", "1.1.2", "HP Equipment Replacement Allowance", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 3, "1.1", "1.1.2", "DGC OH&P on Equipment and Materials", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 4, "1.2", "1.2.3", "General Employee Training Labour", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 5, "1.7", "1.7", "Mixed Waste Disposal", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 6, "1.7", "1.7", "Clean Waste Disposal", "AUD", 1
    V091_EnsureLibraryItem "Model Allowances", 7, "1.5", "1.5.3", "Zero-Cost Waste Volume Carrier", "AUD", 1
End Sub

Private Sub V091_EnsureLibraryItem(ByVal cat As String, ByVal itemID As Long, ByVal wbs As String, ByVal wbsSub As String, ByVal itemName As String, ByVal unitName As String, ByVal baseRate As Double)
    If DCount("*", "tblCostLibrary", "CategoryName=" & V091_Q(cat) & " AND ItemID=" & itemID) > 0 Then Exit Sub
    CurrentDb.Execute "INSERT INTO tblCostLibrary (CategoryName, ItemID, IsActive, WBSCode, WBSSubCode, ItemName, UnitName, BaseUnitRateUSD2009, CreatedAt, UpdatedAt) VALUES (" & _
        V091_Q(cat) & ", " & itemID & ", True, " & V091_Q(wbs) & ", " & V091_Q(wbsSub) & ", " & V091_Q(itemName) & ", " & V091_Q(unitName) & ", " & V091_SqlNum(baseRate) & ", Now(), Now());", dbFailOnError
End Sub

' ============================================================
' QUERIES
' ============================================================

Private Sub V091_CreateQueries()
    V091_SaveQuery "qryV091GeneratedLineEdit", _
        "SELECT GeneratedLineID, JobID, WBSCode, WBSSubCode, Description, LineType, TargetCategoryName, TargetItemID, QuantityPhysical, UnitName, UnitRateAUD, AdjustmentFactor, GeneratedCostAUD, CleanVolM3, LsaVolM3, HazVolM3, MixedVolM3, SourceNote, AppliedToJobLine " & _
        "FROM tblV091GeneratedLines ORDER BY JobID, WBSCode, WBSSubCode, GeneratedLineID;"

    V091_SaveQuery "qryV091Totals", _
        "SELECT JobID, Sum(Nz(GeneratedCostAUD,0)) AS V091GeneratedSubtotalAUD, Sum(Nz(CleanVolM3,0)) AS TotalCleanVolM3, Sum(Nz(LsaVolM3,0)) AS TotalLsaVolM3, Sum(Nz(HazVolM3,0)) AS TotalHazVolM3, Sum(Nz(MixedVolM3,0)) AS TotalMixedVolM3 " & _
        "FROM tblV091GeneratedLines GROUP BY JobID;"

    V091_SaveQuery "qryV091WbsSummary", _
        "SELECT JobID, WBSCode, WBSSubCode, Min(Description) AS ExampleDescription, Sum(Nz(GeneratedCostAUD,0)) AS WbsCostAUD, Sum(Nz(CleanVolM3,0)) AS WbsCleanVolM3, Sum(Nz(LsaVolM3,0)) AS WbsLsaVolM3, Sum(Nz(HazVolM3,0)) AS WbsHazVolM3, Sum(Nz(MixedVolM3,0)) AS WbsMixedVolM3 " & _
        "FROM tblV091GeneratedLines GROUP BY JobID, WBSCode, WBSSubCode ORDER BY JobID, WBSCode, WBSSubCode;"

    V091_SaveQuery "qryV091PortfolioOverview", _
        "SELECT j.JobID, j.JobName, b.BuildingCode, b.FacilityClass, b.FacilityType, b.EstimateBasis, b.TotalAreaM2, b.ProjectDurationDays, b.CrewSize, b.RemovalAdjustmentPct, Nz(v.V091GeneratedSubtotalAUD,0) AS V091GeneratedSubtotalAUD, Nz(g.SubtotalAUD,0) AS FineTuneSubtotalAUD, Nz(g.GrandTotalAUD,0) AS FineTuneGrandTotalAUD, Nz(g.SubtotalAUD,0)-Nz(v.V091GeneratedSubtotalAUD,0) AS DeltaAUD, b.LastGeneratedAt " & _
        "FROM ((tblJobs AS j LEFT JOIN tblV091BuildingInputs AS b ON j.JobID=b.JobID) LEFT JOIN qryV091Totals AS v ON j.JobID=v.JobID) LEFT JOIN qryGrandTotals AS g ON j.JobID=g.JobID ORDER BY j.JobID;"
End Sub

Private Sub V091_SaveQuery(ByVal queryName As String, ByVal sqlText As String)
    On Error Resume Next
    CurrentDb.QueryDefs.Delete queryName
    On Error GoTo 0
    CurrentDb.CreateQueryDef queryName, sqlText
End Sub

' ============================================================
' GENERATION ENGINE
' ============================================================

Public Sub V091_GenerateGenericEstimate(ByVal jobID As String)
    On Error GoTo ErrHandler

    If DCount("*", "tblJobs", "JobID=" & V091_Q(jobID)) = 0 Then Err.Raise vbObjectError + 8210, , "JobID not found: " & jobID
    V091_CreateInputIfMissing jobID

    CurrentDb.Execute "DELETE FROM tblV091GeneratedLines WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError

    V091_GenerateManagementStaff jobID
    V091_GeneratePlanningAndSitePrep jobID
    V091_GenerateDetailedCharacterization jobID
    V091_GenerateAreaActivities jobID
    V091_GenerateWasteDisposal jobID
    V091_GenerateEquipmentConsumablesAndAllowances jobID

    CurrentDb.Execute "UPDATE tblV091BuildingInputs SET LastGeneratedAt=Now(), UpdatedAt=Now() WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError
    Exit Sub

ErrHandler:
    Err.Raise Err.Number, , "V091_GenerateGenericEstimate failed for " & jobID & ": " & Err.Description
End Sub

Private Sub V091_GenerateManagementStaff(ByVal jobID As String)
    Dim d As Double
    Dim portfolioUse As Double
    Dim seniorUse As Double
    Dim wasteUse As Double
    Dim portfolioNumber As Double
    Dim seniorNumber As Double
    Dim wasteNumber As Double

    d = V091_InputDbl(jobID, "ProjectDurationDays", 105)

    ' These are explicit Excel judgement levers. Number and use factor both
    ' exist in the workbook staff table. Defaults preserve the validated v0.8.9
    ' outputs when fields are left blank.
    portfolioUse = V091_InputDbl(jobID, "PortfolioManagerUseFactor", 1)
    seniorUse = V091_InputDbl(jobID, "SeniorPMUseFactor", 0.5)
    wasteUse = V091_InputDbl(jobID, "ProjectManagerUseFactor", 0.5)
    portfolioNumber = V091_InputDbl(jobID, "PortfolioManagerNumber", 1)
    seniorNumber = V091_InputDbl(jobID, "SeniorPMNumber", 1)
    wasteNumber = V091_InputDbl(jobID, "ProjectManagerNumber", 1)

    V091_AddLine jobID, "1.1", "1.1.1", "Portfolio Manager", "Labour", "Labour", 48, d * 8 * portfolioNumber * portfolioUse, "hr", V091_LabourRate(48), 1, 0, 0, 0, 0, "Portfolio manager number and use factor from estimate controls."
    V091_AddLine jobID, "1.1", "1.1.1", "Senior Project Manager / Characterization SME", "Labour", "Labour", 47, d * 10 * seniorNumber * seniorUse, "hr", V091_LabourRate(47), 1, 0, 0, 0, 0, "Senior PM / characterization SME number and use factor from estimate controls."
    V091_AddLine jobID, "1.1", "1.1.1", "Project Manager / Waste Management", "Labour", "Labour", 46, d * 10 * wasteNumber * wasteUse, "hr", V091_LabourRate(46), 1, 0, 0, 0, 0, "Project manager / waste manager number and use factor from estimate controls."
End Sub

Private Sub V091_GeneratePlanningAndSitePrep(ByVal jobID As String)
    Dim crew As Double
    Dim peopleTraining As Double
    Dim costPerHourTraining As Double
    Dim procHours As Double
    Dim qaHours As Double
    Dim siteTasks As Double
    Dim siteHoursPerTask As Double
    Dim portfolioNumber As Double
    Dim seniorNumber As Double
    Dim wasteNumber As Double

    crew = V091_InputDbl(jobID, "CrewSize", 24)
    procHours = V091_InputDbl(jobID, "ProcedureHours", 80)
    qaHours = V091_InputDbl(jobID, "QASafetyHours", 80)
    siteTasks = V091_SitePrepSelectedTaskCount(jobID)
    siteHoursPerTask = V091_InputDbl(jobID, "SitePrepHoursPerTask", 16)
    If siteHoursPerTask <= 0 Then siteHoursPerTask = 16

    portfolioNumber = V091_InputDbl(jobID, "PortfolioManagerNumber", 1)
    seniorNumber = V091_InputDbl(jobID, "SeniorPMNumber", 1)
    wasteNumber = V091_InputDbl(jobID, "ProjectManagerNumber", 1)

    ' Procedure development and QA are explicit duration levers from the Excel template.
    V091_AddLine jobID, "1.2", "1.2.1", "Procedure Development - Project Specialist", "Labour", "Labour", 45, 2 * procHours, "hr", V091_LabourRate(45), 1, 0, 0, 0, 0, "2 Project Specialists x procedure hours."
    V091_AddLine jobID, "1.2", "1.2.1", "Procedure Development - Project Manager", "Labour", "Labour", 46, procHours, "hr", V091_LabourRate(46), 1, 0, 0, 0, 0, "1 Project Manager x procedure hours."

    V091_AddLine jobID, "1.2", "1.2.2", "QA/Safety Documents - Project Specialist", "Labour", "Labour", 45, 2 * qaHours, "hr", V091_LabourRate(45), 1, 0, 0, 0, 0, "2 Project Specialists x QA/safety hours."
    V091_AddLine jobID, "1.2", "1.2.2", "QA/Safety Documents - Project Manager", "Labour", "Labour", 46, qaHours, "hr", V091_LabourRate(46), 1, 0, 0, 0, 0, "1 Project Manager x QA/safety hours."

    V091_AddLine jobID, "1.2", "1.2.3", "Site Mobilization", "Labour", "Labour", 45, crew * 40, "hr", V091_LabourRate(45), 1, 0, 0, 0, 0, "40 hr x decom crew size."

    If V091_InputBool(jobID, "IncludeTraining", True) Then
        peopleTraining = crew + portfolioNumber + seniorNumber + wasteNumber
        costPerHourTraining = portfolioNumber * V091_LabourRate(48) + seniorNumber * V091_LabourRate(47) + wasteNumber * V091_LabourRate(46) + crew * V091_LabourRate(45)
        V091_AddAllowance jobID, "1.2", "1.2.3", "General Employee Training Labour", 4, V091_TrainingLabourCost(peopleTraining, costPerHourTraining), "Excel parity v0.9.1: WBS 1.2.3 includes GET labour only; direct medical/test costs are shown in the Excel subtable but are not rolled into the WBS total."
    End If

    If V091_InputBool(jobID, "IncludeSitePrep", True) And siteTasks > 0 Then
        V091_AddLine jobID, "1.3", "1.3.2", "Site Prep - Project Specialists", "Labour", "Labour", 45, 4 * siteHoursPerTask * siteTasks, "hr", V091_LabourRate(45), 1, 0, 0, 0, 0, "Selected site prep tasks x hours/task x 4 specialists."
        V091_AddLine jobID, "1.3", "1.3.2", "Site Prep - Project Manager", "Labour", "Labour", 46, siteHoursPerTask * siteTasks, "hr", V091_LabourRate(46), 1, 0, 0, 0, 0, "Selected site prep tasks x hours/task x 1 PM."
    End If
End Sub

Private Function V091_SitePrepSelectedTaskCount(ByVal jobID As String) As Double
    Dim selectedCount As Double
    Dim legacyCount As Double

    selectedCount = 0
    If V091_InputBool(jobID, "SitePrepInitialSurvey", False) Then selectedCount = selectedCount + 1
    If V091_InputBool(jobID, "SitePrepBoundariesHepa", False) Then selectedCount = selectedCount + 1
    If V091_InputBool(jobID, "SitePrepStagingArea", False) Then selectedCount = selectedCount + 1
    If V091_InputBool(jobID, "SitePrepRadSegregation", False) Then selectedCount = selectedCount + 1
    If V091_InputBool(jobID, "SitePrepElectricalIsolation", False) Then selectedCount = selectedCount + 1
    If V091_InputBool(jobID, "SitePrepPipingIsolation", False) Then selectedCount = selectedCount + 1

    If selectedCount > 0 Then
        V091_SitePrepSelectedTaskCount = selectedCount
    Else
        legacyCount = V091_InputDbl(jobID, "SitePrepTaskCount", 0)
        V091_SitePrepSelectedTaskCount = legacyCount
    End If
End Function

Private Function V091_TrainingLabourCost(ByVal people As Double, ByVal costPerHour As Double) As Double
    ' Excel parity v0.9.1:
    ' The Excel training subtable shows direct medical/test costs plus labour costs,
    ' but the WBS 1.2.3 roll-up carries the General Employee Training labour component.
    ' Therefore this returns only the 55-hour labour component, not the direct test-cost component.
    Dim labourCost As Double
    labourCost = (4 * costPerHour) + (1 * costPerHour) + (1 * costPerHour) + (1 * costPerHour) + (8 * costPerHour) + (40 * costPerHour)
    V091_TrainingLabourCost = labourCost
End Function

Private Sub V091_GenerateDetailedCharacterization(ByVal jobID As String)
    Dim specCount As Double
    Dim pmCount As Double
    Dim hoursEach As Double

    If Not V091_InputBool(jobID, "IncludeDetailedCharacterization", True) Then Exit Sub

    specCount = V091_InputDbl(jobID, "CharacterizationSpecialistCount", 8)
    pmCount = V091_InputDbl(jobID, "CharacterizationPMCount", 2)
    hoursEach = V091_InputDbl(jobID, "CharacterizationHoursPerPerson", 320)

    V091_AddLine jobID, "1.4", "1.4", "Detailed Characterization - Project Specialists", "Labour", "Labour", 45, specCount * hoursEach, "hr", V091_LabourRate(45), 1, 0, 0, 0, 0, "Characterization specialist count and hours from estimate controls."
    V091_AddLine jobID, "1.4", "1.4", "Detailed Characterization - Project Managers", "Labour", "Labour", 46, pmCount * hoursEach, "hr", V091_LabourRate(46), 1, 0, 0, 0, 0, "Characterization PM count and hours from estimate controls."
End Sub

Private Sub V091_GenerateAreaActivities(ByVal jobID As String)
    Dim rs As DAO.Recordset
    Dim basis As String, isRad As Boolean
    Dim area As Double, footprint As Double, remMult As Double, crewFactor As Double
    Dim qty As Double, cleanV As Double, lsaV As Double, hazV As Double, mixedV As Double
    Dim appliedFactor As Double, rateAUD As Double

    basis = V091_InputText(jobID, "EstimateBasis", "Building / Structure D&D")
    isRad = V091_InputBool(jobID, "IsRadiological", True)
    area = V091_InputDbl(jobID, "TotalAreaM2", 0)
    footprint = V091_InputDbl(jobID, "FootprintAreaM2", area)
    If footprint <= 0 Then footprint = area
    remMult = 1 + V091_InputDbl(jobID, "RemovalAdjustmentPct", 0) / 100
    If remMult < 0 Then remMult = 0

    ' v0.9.1 formula parity cleanup:
    ' The Excel building sheets calculate the activity rate as:
    '   Summary 2026 base rate x (crew size / 12)
    ' The class/type adjustment is then represented by RemovalAdjustmentPct,
    ' applied as (1 + RemovalAdjustmentPct) only to rows that use the removal adjustment.
    crewFactor = V091_InputDbl(jobID, "CrewSize", 12) / 12
    If crewFactor <= 0 Then crewFactor = 1

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblV091AreaActivityTemplate WHERE DefaultEnabled=True AND EstimateBasis=" & V091_Q(basis) & " ORDER BY WBSCode, WBSSubCode, TemplateID;", dbOpenSnapshot)
    Do While Not rs.EOF
        If (Not Nz(rs!RequiresRadiological, False)) Or isRad Then
            qty = V091_ActivityQty(jobID, Nz(rs!quantitySource, ""), area, footprint)
            cleanV = qty * CDbl(Nz(rs!CleanVolRateM3, 0))
            lsaV = qty * CDbl(Nz(rs!LsaVolRateM3, 0))
            hazV = qty * CDbl(Nz(rs!HazVolRateM3, 0))
            mixedV = qty * CDbl(Nz(rs!MixedVolRateM3, 0))
            If qty > 0 Or cleanV <> 0 Or lsaV <> 0 Or hazV <> 0 Or mixedV <> 0 Then
                rateAUD = CDbl(Nz(rs!unitRateAUD, 0)) * crewFactor
                If basis = "Building / Structure D&D" And Nz(rs!Description, "") = "Remove hazardous material" Then
                    If Not V091_InputBool(jobID, "IncludeHotCell", False) Then rateAUD = 0
                End If
                appliedFactor = 1
                If Nz(rs!ApplyRemovalAdjustment, False) Then appliedFactor = remMult
                V091_AddLine jobID, Nz(rs!wbsCode, ""), Nz(rs!WBSSubCode, ""), Nz(rs!Description, ""), "AreaActivity", Nz(rs!categoryName, ""), CLng(Nz(rs!itemID, 0)), qty, Nz(rs!unitName, ""), rateAUD, appliedFactor, cleanV, lsaV, hazV, mixedV, Nz(rs!notes, "") & " Rate = Summary 2026 base rate x crew size/12."
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function V091_ActivityQty(ByVal jobID As String, ByVal quantitySource As String, ByVal area As Double, ByVal footprint As Double) As Double
    Dim removalDepth As Double, backfillDepth As Double, pctClean As Double, pctCont As Double
    removalDepth = V091_InputDbl(jobID, "RemovalDepthM", 1)
    backfillDepth = V091_InputDbl(jobID, "BackfillDepthM", 1)
    pctClean = V091_InputDbl(jobID, "PercentClean", 50) / 100
    pctCont = V091_InputDbl(jobID, "PercentContaminated", 50) / 100

    Select Case quantitySource
        Case "TotalAreaM2"
            V091_ActivityQty = area
        Case "AsbestosPipeLengthM"
            V091_ActivityQty = V091_InputDbl(jobID, "AsbestosPipeLengthM", 0)
        Case "AsbestosTileAreaM2"
            V091_ActivityQty = V091_InputDbl(jobID, "AsbestosTileAreaM2", 0)
        Case "GradeAreaM2"
            V091_ActivityQty = footprint * 2
        Case "BackfillM3"
            V091_ActivityQty = footprint * backfillDepth
        Case "CleanSoilVolumeM3"
            V091_ActivityQty = area * removalDepth * pctClean
        Case "ContaminatedSoilVolumeM3"
            V091_ActivityQty = area * removalDepth * pctCont
        Case "HotCellAreaM2"
            If V091_InputBool(jobID, "IncludeHotCell", False) Then
                V091_ActivityQty = V091_InputDbl(jobID, "HotCellAreaM2", 0)
                If V091_ActivityQty <= 0 Then V091_ActivityQty = area
            Else
                V091_ActivityQty = 0
            End If
        Case "Zero"
            V091_ActivityQty = 0
        Case Else
            V091_ActivityQty = 0
    End Select
End Function

Private Sub V091_GenerateWasteDisposal(ByVal jobID As String)
    Dim cleanM3 As Double, lsaM3 As Double, hazM3 As Double, mixedM3 As Double
    Dim density As Double, m3ToCuft As Double, lbPerTon As Double
    Dim cleanTons As Double, hazTons As Double, mixedTons As Double

    cleanM3 = Nz(DSum("CleanVolM3", "tblV091GeneratedLines", "JobID=" & V091_Q(jobID)), 0)
    lsaM3 = Nz(DSum("LsaVolM3", "tblV091GeneratedLines", "JobID=" & V091_Q(jobID)), 0)
    hazM3 = Nz(DSum("HazVolM3", "tblV091GeneratedLines", "JobID=" & V091_Q(jobID)), 0)
    mixedM3 = Nz(DSum("MixedVolM3", "tblV091GeneratedLines", "JobID=" & V091_Q(jobID)), 0)

    density = V091_SettingDbl("WasteDensityLbPerCuft", 100)
    m3ToCuft = V091_SettingDbl("M3ToCuft", 35.3147)
    lbPerTon = V091_SettingDbl("LbPerWasteTon", 2000)

    cleanTons = cleanM3 * m3ToCuft * density / lbPerTon
    hazTons = hazM3 * m3ToCuft * density / lbPerTon
    mixedTons = mixedM3 * m3ToCuft * density / lbPerTon

    V091_AddLine jobID, "1.7", "1.7", "Clean Waste Disposal", "Waste", "Model Allowances", 6, cleanTons, "ton", V091_SettingDbl("IndustrialWasteRate", 300), 1, 0, 0, 0, 0, "Waste derived from generated clean volume."
    V091_AddLine jobID, "1.7", "1.7", "Hazardous Waste Disposal", "Waste", "Waste Disposal", 37, hazTons, "ton", V091_SettingDbl("HazardousWasteRate", 500), 1, 0, 0, 0, 0, "Waste derived from generated hazardous volume."
    V091_AddLine jobID, "1.7", "1.7", "Mixed Waste Disposal", "Waste", "Model Allowances", 5, mixedTons, "ton", V091_SettingDbl("HazardousWasteRate", 500), 1, 0, 0, 0, 0, "Waste derived from generated mixed volume."
    V091_AddLine jobID, "1.7", "1.7", "LSA/LLW Stored Waste", "Waste", "Waste Disposal", 38, lsaM3, "m3", 0, 1, 0, 0, 0, 0, "LSA/LLW volume tracked; cost rate is zero in template."
End Sub

Private Sub V091_GenerateEquipmentConsumablesAndAllowances(ByVal jobID As String)
    ' Excel formula parity for 1.1.2 Equipment, Materials & Consumables:
    '   S8 = (H46 + M58 + K60 + K62 + K64) * Inputs!C37
    ' where:
    '   H46 = escalated equipment subtotal before AUD conversion
    '   M58 = escalated consumables subtotal before AUD conversion
    '   K60 = Small Tools = 2% of (T6 - T8)
    '   K62 = HP Equipment Replacement = 5% of (T6 - T8)
    '   K64 = DGC OH&P = 8% of (H46 + M58 + K60 + K62 + any extra basis, usually zero)
    '
    ' To keep generated lines in AUD, equipment/consumable rates and allowance costs are multiplied
    ' by the workbook exchange factor here, but the allowance base is still calculated pre-FX.

    Dim facilityClass As String
    Dim rs As DAO.Recordset
    Dim qty As Double, rateUSD As Double, rateAUD As Double, durationDays As Double, crew As Double
    Dim equipCostPreFx As Double, consCostPreFx As Double, activityCostBase As Double
    Dim smallToolsPreFx As Double, hpReplacePreFx As Double, ohpPreFx As Double
    Dim fx As Double

    facilityClass = V091_InputText(jobID, "FacilityClass", "B1")
    durationDays = V091_InputDbl(jobID, "WorkDays", V091_InputDbl(jobID, "ProjectDurationDays", 105))
    crew = V091_InputDbl(jobID, "CrewSize", 24)
    fx = V091_SettingDbl("UsdToAudExchangeRate", 1.54)

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblV091EquipmentTemplate WHERE FacilityClass=" & V091_Q(facilityClass) & " AND DefaultEnabled=True ORDER BY ItemID;", dbOpenSnapshot)
    Do While Not rs.EOF
        qty = CDbl(Nz(rs!quantity, 0))
        If qty > 0 Then
            rateUSD = V091_BaseRate("Equipment", CLng(rs!itemID)) * V091_SettingDbl("EquipmentEscalation2009To2026", 1.58014)
            rateAUD = rateUSD * fx
            equipCostPreFx = equipCostPreFx + (qty * rateUSD)
            V091_AddLine jobID, "1.1", "1.1.2", Nz(rs!itemName, "Equipment"), "Equipment", "Equipment", CLng(rs!itemID), qty, "item", rateAUD, 1, 0, 0, 0, 0, "Class equipment template quantity. Excel parity: escalated equipment cost converted by Inputs!C37."
        End If
        rs.MoveNext
    Loop
    rs.Close

    If V091_InputBool(jobID, "IncludeConsumables", True) Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblV091ConsumableTemplate WHERE FacilityClass=" & V091_Q(facilityClass) & " AND DefaultEnabled=True ORDER BY ItemID;", dbOpenSnapshot)
        Do While Not rs.EOF
            qty = V091_ConsumableQty(jobID, CDbl(Nz(rs!useRate, 0)), Nz(rs!DurationBasis, "man-day"), durationDays, crew)
            If qty > 0 Then
                rateUSD = V091_BaseRate("Consumables", CLng(rs!itemID)) * V091_SettingDbl("EquipmentEscalation2009To2026", 1.58014)
                rateAUD = rateUSD * fx
                consCostPreFx = consCostPreFx + (qty * rateUSD)
                V091_AddLine jobID, "1.1", "1.1.2", Nz(rs!itemName, "Consumable"), "Consumable", "Consumables", CLng(rs!itemID), qty, Nz(rs!DurationBasis, "unit"), rateAUD, 1, 0, 0, 0, 0, "Class consumable template quantity. Excel parity: escalated consumable cost converted by Inputs!C37."
            End If
            rs.MoveNext
        Loop
        rs.Close
    End If

    activityCostBase = V091_ActivityCostBase(jobID)

    smallToolsPreFx = activityCostBase * V091_SettingDbl("SmallToolsPctOfActivityLabor", 0.02)
    hpReplacePreFx = activityCostBase * V091_SettingDbl("HpEquipmentReplacementPctOfActivityLabor", 0.05)
    ohpPreFx = (equipCostPreFx + consCostPreFx + smallToolsPreFx + hpReplacePreFx) * V091_SettingDbl("EquipmentOverheadPct", 0.08)

    V091_AddAllowance jobID, "1.1", "1.1.2", "Small Tools - 2% of activity cost base", 1, smallToolsPreFx * fx, "Generic BXX allowance. Excel parity: K60 included in S8 and converted by Inputs!C37."
    V091_AddAllowance jobID, "1.1", "1.1.2", "HP Equipment Replacement - 5% of activity cost base", 2, hpReplacePreFx * fx, "Generic BXX allowance. Excel parity: K62 included in S8 and converted by Inputs!C37."
    V091_AddAllowance jobID, "1.1", "1.1.2", "DGC OH&P on equipment and materials", 3, ohpPreFx * fx, "Generic BXX allowance. Excel parity: K64 included in S8 and converted by Inputs!C37."
End Sub

Private Function V091_ConsumableQty(ByVal jobID As String, ByVal useRate As Double, ByVal basis As String, ByVal days As Double, ByVal crew As Double) As Double
    Select Case basis
        Case "man-day"
            V091_ConsumableQty = useRate * days * crew
        Case "man-month"
            V091_ConsumableQty = useRate * V091_InputDbl(jobID, "ConsumableMonths", 3) * crew
        Case "man-year"
            V091_ConsumableQty = useRate * V091_InputDbl(jobID, "DosimeterYears", 0.25) * crew
        Case "man-year-bioassay"
            V091_ConsumableQty = useRate * V091_InputDbl(jobID, "BioassayYears", 0.5) * crew
        Case Else
            V091_ConsumableQty = 0
    End Select
End Function

Private Function V091_ActivityCostBase(ByVal jobID As String) As Double
    Dim whereText As String
    whereText = "JobID=" & V091_Q(jobID) & " AND WBSSubCode<>'1.1.1' AND WBSSubCode<>'1.1.2' AND WBSSubCode<>'1.7'"
    V091_ActivityCostBase = Nz(DSum("GeneratedCostAUD", "tblV091GeneratedLines", whereText), 0)
End Function

Private Sub V091_AddAllowance(ByVal jobID As String, ByVal wbsCode As String, ByVal wbsSub As String, ByVal descText As String, ByVal allowanceItemID As Long, ByVal costAUD As Double, ByVal note As String)
    If costAUD <= 0 Then Exit Sub
    V091_AddLine jobID, wbsCode, wbsSub, descText, "Allowance", "Model Allowances", allowanceItemID, costAUD, "AUD", 1, 1, 0, 0, 0, 0, note
End Sub

Private Sub V091_AddLine(ByVal jobID As String, ByVal wbsCode As String, ByVal wbsSub As String, ByVal descText As String, ByVal lineType As String, ByVal targetCat As String, ByVal targetItemID As Long, ByVal qty As Double, ByVal unitName As String, ByVal unitRateAUD As Double, ByVal adjustmentFactor As Double, ByVal cleanV As Double, ByVal lsaV As Double, ByVal hazV As Double, ByVal mixedV As Double, ByVal sourceNote As String)
    Dim amount As Double
    If adjustmentFactor < 0 Then adjustmentFactor = 0
    amount = qty * unitRateAUD * adjustmentFactor
    If amount = 0 And cleanV = 0 And lsaV = 0 And hazV = 0 And mixedV = 0 Then Exit Sub

    CurrentDb.Execute "INSERT INTO tblV091GeneratedLines " & _
        "(JobID, WBSCode, WBSSubCode, Description, LineType, TargetCategoryName, TargetItemID, QuantityPhysical, UnitName, UnitRateAUD, AdjustmentFactor, GeneratedCostAUD, CleanVolM3, LsaVolM3, HazVolM3, MixedVolM3, SourceNote, AppliedToJobLine, CreatedAt) VALUES (" & _
        V091_Q(jobID) & ", " & V091_Q(wbsCode) & ", " & V091_Q(wbsSub) & ", " & V091_Q(descText) & ", " & V091_Q(lineType) & ", " & V091_Q(targetCat) & ", " & targetItemID & ", " & V091_SqlNum(qty) & ", " & V091_Q(unitName) & ", " & V091_SqlCurrency(unitRateAUD) & ", " & V091_SqlNum(adjustmentFactor) & ", " & V091_SqlCurrency(amount) & ", " & V091_SqlNum(cleanV) & ", " & V091_SqlNum(lsaV) & ", " & V091_SqlNum(hazV) & ", " & V091_SqlNum(mixedV) & ", " & V091_Q(sourceNote) & ", False, Now());", dbFailOnError
End Sub

' ============================================================
' APPLY GENERATED LINES TO EXISTING JOB LINES
' ============================================================

Public Sub V091_ApplyGeneratedToJobLines(ByVal jobID As String)
    On Error GoTo ErrHandler
    Dim rs As DAO.Recordset
    Dim cat As String, itemID As Long, qty As Double, cost As Double, desiredRateAUD As Double, baseForV062 As Double

    If DCount("*", "tblV091GeneratedLines", "JobID=" & V091_Q(jobID)) = 0 Then V091_GenerateGenericEstimate jobID

    V091_BackupJobLines jobID
    V091_EnsureJobHasAllActiveLibraryLines jobID

    CurrentDb.Execute "UPDATE tblJobLines SET IncludeItem=False, Quantity=0, V091GeneratedQuantity=Null, V091GeneratedCostAUD=Null, V091IsGenerated=False WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError

    Set rs = CurrentDb.OpenRecordset("SELECT TargetCategoryName, TargetItemID, Sum(Nz(QuantityPhysical,0)) AS TotalQty, Sum(Nz(GeneratedCostAUD,0)) AS TotalCostAUD FROM tblV091GeneratedLines WHERE JobID=" & V091_Q(jobID) & " AND TargetCategoryName Is Not Null GROUP BY TargetCategoryName, TargetItemID;", dbOpenSnapshot)
    Do While Not rs.EOF
        cat = Nz(rs!TargetCategoryName, "")
        itemID = CLng(Nz(rs!TargetItemID, 0))
        qty = CDbl(Nz(rs!TotalQty, 0))
        cost = CDbl(Nz(rs!TotalCostAUD, 0))

        If cat <> "" And itemID > 0 And cost <> 0 Then
            If qty <= 0 Then qty = cost
            desiredRateAUD = cost / qty
            baseForV062 = desiredRateAUD / V091_V062Multiplier()
            CurrentDb.Execute "UPDATE tblJobLines SET IncludeItem=True, Quantity=" & V091_SqlNum(qty) & _
                ", V091GeneratedQuantity=" & V091_SqlNum(qty) & _
                ", V091GeneratedCostAUD=" & V091_SqlCurrency(cost) & _
                ", V091IsGenerated=True, V091LastGeneratedAt=Now(), " & _
                "V091OriginalBaseRateUSD2009=IIf(V091OriginalBaseRateUSD2009 Is Null, BaseUnitRateUSD2009, V091OriginalBaseRateUSD2009), " & _
                "BaseUnitRateUSD2009=" & V091_SqlCurrency(baseForV062) & _
                " WHERE JobID=" & V091_Q(jobID) & " AND CategoryName=" & V091_Q(cat) & " AND ItemID=" & itemID & ";", dbFailOnError
        End If
        rs.MoveNext
    Loop
    rs.Close

    CurrentDb.Execute "UPDATE tblV091GeneratedLines SET AppliedToJobLine=True WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, , "V091_ApplyGeneratedToJobLines failed: " & Err.Description
End Sub

Private Sub V091_BackupJobLines(ByVal jobID As String)
    CurrentDb.Execute "INSERT INTO tblV091JobLineBackup (BackupAt, JobID, JobLineID, CategoryName, ItemID, IncludeItem, Quantity, BaseUnitRateUSD2009) " & _
        "SELECT Now(), JobID, JobLineID, CategoryName, ItemID, IncludeItem, Quantity, BaseUnitRateUSD2009 FROM tblJobLines WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError
End Sub

Private Sub V091_EnsureJobHasAllActiveLibraryLines(ByVal jobID As String)
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT LibraryID, CategoryName, ItemID, WBSCode, WBSSubCode, ItemName, UnitName, BaseUnitRateUSD2009 FROM tblCostLibrary WHERE IsActive=True ORDER BY CategoryName, ItemID;", dbOpenSnapshot)
    Do While Not rs.EOF
        If DCount("*", "tblJobLines", "JobID=" & V091_Q(jobID) & " AND CategoryName=" & V091_Q(Nz(rs!CategoryName, "")) & " AND ItemID=" & CLng(Nz(rs!itemID, 0))) = 0 Then
            CurrentDb.Execute "INSERT INTO tblJobLines (JobID, LibraryID, CategoryName, ItemID, IncludeItem, WBSCode, WBSSubCode, ItemName, Quantity, UnitName, BaseUnitRateUSD2009) VALUES (" & _
                V091_Q(jobID) & ", " & CLng(Nz(rs!LibraryID, 0)) & ", " & V091_Q(Nz(rs!CategoryName, "")) & ", " & CLng(Nz(rs!itemID, 0)) & ", False, " & V091_Q(Nz(rs!wbsCode, "")) & ", " & V091_Q(Nz(rs!WBSSubCode, "")) & ", " & V091_Q(Nz(rs!itemName, "")) & ", 0, " & V091_Q(Nz(rs!UnitName, "")) & ", " & V091_SqlCurrency(CDbl(Nz(rs!BaseUnitRateUSD2009, 0))) & ");", dbFailOnError
        End If
        rs.MoveNext
    Loop
    rs.Close
End Sub

' ============================================================
' UI FUNCTIONS
' ============================================================

Public Function V091_UI_NewJob() As Boolean
    On Error GoTo ErrHandler
    Dim jobID As String, jobName As String, preparedBy As String
    jobID = Trim(InputBox("Enter Job ID / Building ID.", "New v0.9.1 Estimate", ""))
    If jobID = "" Then Exit Function
    jobName = Trim(InputBox("Enter job/building name.", "New v0.9.1 Estimate", ""))
    If jobName = "" Then jobName = "Building " & jobID
    preparedBy = Trim(InputBox("Prepared by:", "New v0.9.1 Estimate", ""))

    If DCount("*", "tblJobs", "JobID=" & V091_Q(jobID)) = 0 Then
        V062_CreateNewJobFromLibrary jobID, jobName, preparedBy
    End If
    V091_CreateInputIfMissing jobID
    Forms("frmV091GenericEstimate").Requery
    V091_MoveFormToJob "frmV091GenericEstimate", jobID
    V091_UI_NewJob = True
    Exit Function
ErrHandler:
    MsgBox "Could not create v0.9.1 job:" & vbCrLf & Err.Description, vbExclamation, "v0.9.1 New Job Error"
    V091_UI_NewJob = False
End Function

Public Function V091_UI_LoadSelectedJob() As Boolean
    On Error GoTo ErrHandler
    Dim jobID As String
    jobID = Nz(Forms("frmV091GenericEstimate").Controls("cboJobSelector").value, "")
    If jobID = "" Then Exit Function
    V091_CreateInputIfMissing jobID
    Forms("frmV091GenericEstimate").Requery
    V091_MoveFormToJob "frmV091GenericEstimate", jobID
    V091_UI_Refresh
    V091_UI_LoadSelectedJob = True
    Exit Function
ErrHandler:
    MsgBox "Could not load job:" & vbCrLf & Err.Description, vbExclamation, "v0.9.1 Load Error"
    V091_UI_LoadSelectedJob = False
End Function

Public Function V091_UI_Generate() As Boolean
    On Error GoTo ErrHandler
    Dim jobID As String
    jobID = V091_CurrentFormJobID()
    If jobID = "" Then Exit Function
    If Forms("frmV091GenericEstimate").Dirty Then Forms("frmV091GenericEstimate").Dirty = False
    V091_GenerateGenericEstimate jobID
    V091_UI_Refresh
    MsgBox "v0.9.1 generic estimate generated for " & jobID & ".", vbInformation, "Generated"
    V091_UI_Generate = True
    Exit Function
ErrHandler:
    MsgBox "Could not generate v0.9.1 estimate:" & vbCrLf & Err.Description, vbExclamation, "v0.9.1 Generate Error"
    V091_UI_Generate = False
End Function

Public Function V091_UI_Apply() As Boolean
    On Error GoTo ErrHandler
    Dim jobID As String
    jobID = V091_CurrentFormJobID()
    If jobID = "" Then Exit Function
    If MsgBox("Apply v0.9.1 generated lines to the detailed job line copy for " & jobID & "?" & vbCrLf & vbCrLf & _
              "This changes only tblJobLines for this job. tblCostLibrary is not overwritten.", vbQuestion + vbYesNo, "Apply v0.9.1") <> vbYes Then Exit Function
    V091_ApplyGeneratedToJobLines jobID
    V091_UI_Refresh
    MsgBox "v0.9.1 generated scope applied to detailed job lines.", vbInformation, "Applied"
    V091_UI_Apply = True
    Exit Function
ErrHandler:
    MsgBox "Could not apply v0.9.1 estimate:" & vbCrLf & Err.Description, vbExclamation, "v0.9.1 Apply Error"
    V091_UI_Apply = False
End Function

Public Function V091_UI_OpenFineTune() As Boolean
    On Error GoTo ErrHandler
    Dim jobID As String
    jobID = V091_CurrentFormJobID()
    If jobID = "" Then Exit Function
    DoCmd.OpenForm "frmJobEstimate"
    V091_MoveFormToJob "frmJobEstimate", jobID
    On Error Resume Next
    V062_UI_Refresh
    On Error GoTo 0
    V091_UI_OpenFineTune = True
    Exit Function
ErrHandler:
    MsgBox "Could not open fine tuning:" & vbCrLf & Err.Description, vbExclamation, "Open Fine Tune Error"
    V091_UI_OpenFineTune = False
End Function

Public Function V091_UI_OpenPortfolio() As Boolean
    On Error Resume Next
    DoCmd.Close acForm, "frmV091PortfolioOverview", acSaveNo
    On Error GoTo 0
    DoCmd.OpenForm "frmV091PortfolioOverview"
    V091_UI_OpenPortfolio = True
End Function

Public Function V091_UI_Refresh() As Boolean
    On Error Resume Next
    If CurrentProject.AllForms("frmV091GenericEstimate").IsLoaded Then
        Forms("frmV091GenericEstimate").Controls("cboJobSelector").Requery
        Forms("frmV091GenericEstimate").Controls("cboJobSelector").value = V091_CurrentFormJobID()
        Forms("frmV091GenericEstimate").Controls("subV091GeneratedLines").Form.Requery
        Forms("frmV091GenericEstimate").Recalc
        Forms("frmV091GenericEstimate").Repaint
    End If
    V091_UI_Refresh = True
End Function

Public Function V091_UI_ClassChanged() As Boolean
    On Error GoTo ErrHandler

    Dim cls As String
    Dim jobID As String

    cls = Nz(Forms("frmV091GenericEstimate").Controls("cboClass").value, "")
    jobID = V091_CurrentFormJobID()
    If cls = "" Or jobID = "" Then Exit Function

    If Forms("frmV091GenericEstimate").Dirty Then Forms("frmV091GenericEstimate").Dirty = False

    CurrentDb.Execute "UPDATE tblV091BuildingInputs SET FacilityType=" & V091_Q(V091_FacilityTypeFromClass(cls)) & _
        ", IsRadiological=" & IIf(V091_ClassIsRadiological(cls), "True", "False") & _
        ", RemovalAdjustmentPct=" & V091_SqlNum(V091_DefaultRemovalAdjustmentFromClass(cls)) & _
        ", UpdatedAt=Now() WHERE JobID=" & V091_Q(jobID) & ";", dbFailOnError

    Forms("frmV091GenericEstimate").Requery
    V091_MoveFormToJob "frmV091GenericEstimate", jobID
    V091_UI_ClassChanged = True
    Exit Function

ErrHandler:
    MsgBox "Could not apply class defaults:" & vbCrLf & Err.Description, vbExclamation, "Class Defaults Error"
    V091_UI_ClassChanged = False
End Function

Private Function V091_FacilityTypeFromClass(ByVal cls As String) As String
    Select Case cls
        Case "B1": V091_FacilityTypeFromClass = "Lab/Rad"
        Case "C1": V091_FacilityTypeFromClass = "Industrial/Rad"
        Case "D1": V091_FacilityTypeFromClass = "Non-Indust/Rad"
        Case "D2": V091_FacilityTypeFromClass = "Non-Indust/Non-Rad"
        Case Else: V091_FacilityTypeFromClass = ""
    End Select
End Function

Private Function V091_DefaultRemovalAdjustmentFromClass(ByVal cls As String) As Double
    Select Case cls
        Case "B1": V091_DefaultRemovalAdjustmentFromClass = 50
        Case "C1": V091_DefaultRemovalAdjustmentFromClass = 0
        Case "D1", "D2": V091_DefaultRemovalAdjustmentFromClass = -50
        Case Else: V091_DefaultRemovalAdjustmentFromClass = 0
    End Select
End Function

Private Function V091_ClassIsRadiological(ByVal cls As String) As Boolean
    V091_ClassIsRadiological = (cls = "B1" Or cls = "C1" Or cls = "D1")
End Function

Private Function V091_CurrentFormJobID() As String
    On Error GoTo Fallback
    V091_CurrentFormJobID = Nz(Forms("frmV091GenericEstimate").Controls("txtJobID").value, "")
    Exit Function
Fallback:
    V091_CurrentFormJobID = ""
End Function

' ============================================================
' FORMS
' ============================================================

Private Sub V091_CreateGeneratedLinesSubform()
    Dim frm As Form, oldName As String
    Set frm = CreateForm
    oldName = frm.Name
    With frm
        .RecordSource = "qryV091GeneratedLineEdit"
        .Caption = "v0.9.1 Generated Lines"
        .DefaultView = 2
        .ViewsAllowed = 2
        .AllowEdits = False
        .AllowAdditions = False
        .AllowDeletions = False
        .NavigationButtons = True
        .RecordSelectors = True
    End With
    V091_AddDatasheetText oldName, "WBS", "WBSSubCode", 0, 0, 850, True
    V091_AddDatasheetText oldName, "Description", "Description", 850, 0, 3600, True
    V091_AddDatasheetText oldName, "Type", "LineType", 4450, 0, 1100, True
    V091_AddDatasheetText oldName, "TargetCat", "TargetCategoryName", 5550, 0, 1500, True
    V091_AddDatasheetText oldName, "TargetID", "TargetItemID", 7050, 0, 800, True
    V091_AddDatasheetText oldName, "Qty", "QuantityPhysical", 7850, 0, 1000, True, "0.00"
    V091_AddDatasheetText oldName, "Unit", "UnitName", 8850, 0, 700, True
    V091_AddDatasheetText oldName, "RateAUD", "UnitRateAUD", 9550, 0, 1200, True, "Currency"
    V091_AddDatasheetText oldName, "Factor", "AdjustmentFactor", 10750, 0, 800, True, "0.00"
    V091_AddDatasheetText oldName, "CostAUD", "GeneratedCostAUD", 11550, 0, 1400, True, "Currency"
    V091_AddDatasheetText oldName, "CleanM3", "CleanVolM3", 12950, 0, 900, True, "0.00"
    V091_AddDatasheetText oldName, "LSA_M3", "LsaVolM3", 13850, 0, 900, True, "0.00"
    V091_AddDatasheetText oldName, "HazM3", "HazVolM3", 14750, 0, 900, True, "0.00"
    V091_AddDatasheetText oldName, "MixedM3", "MixedVolM3", 15650, 0, 900, True, "0.00"
    V091_AddDatasheetText oldName, "Source", "SourceNote", 16550, 0, 5200, True
    DoCmd.Save acForm, oldName
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename "frmV091GeneratedLinesSubform", acForm, oldName
End Sub

Private Sub V091_CreateEstimateForm()
    Dim frm As Form, oldName As String, ctl As Control
    Set frm = CreateForm
    oldName = frm.Name
    With frm
        .RecordSource = "tblV091BuildingInputs"
        .Caption = "v0.9.1 Generic BXX Estimate"
        .DefaultView = 0
        .ViewsAllowed = 1
        .AllowEdits = True
        .AllowAdditions = True
        .AllowDeletions = False
        .NavigationButtons = True
        .RecordSelectors = False
        .ScrollBars = 2
        .Width = 22500
        .OnCurrent = "=V091_UI_Refresh()"
    End With

    Set ctl = CreateControl(oldName, acLabel, acDetail, , , 360, 240, 14000, 420)
    ctl.Name = "lblTitle": ctl.Caption = "v0.9.1 Generic BXX Cost Estimate Generator": ctl.FontSize = 18: ctl.FontBold = True

    Set ctl = CreateControl(oldName, acLabel, acDetail, , , 360, 720, 19000, 300)
    ctl.Name = "lblSubtitle": ctl.Caption = "Universal entry point. Enter new building inputs, select estimate basis, generate WBS lines, then apply to detailed job-line screen."

    V091_AddLabel oldName, "lblLoad", "Load job", 360, 1260, 1200, 300
    Set ctl = CreateControl(oldName, acComboBox, acDetail, , , 1560, 1200, 4800, 360)
    ctl.Name = "cboJobSelector": ctl.RowSourceType = "Table/Query": ctl.RowSource = "SELECT JobID, JobName FROM tblJobs ORDER BY JobID;": ctl.ColumnCount = 2: ctl.BoundColumn = 1: ctl.ColumnWidths = "1800;3000": ctl.LimitToList = True: ctl.AfterUpdate = "=V091_UI_LoadSelectedJob()"

    V091_AddButton oldName, "cmdNew", "New / Attach Job", 6660, 1170, 2200, 450, "=V091_UI_NewJob()"
    V091_AddButton oldName, "cmdGenerate", "Generate v0.9.1", 9060, 1170, 2200, 450, "=V091_UI_Generate()"
    V091_AddButton oldName, "cmdApply", "Apply to Job Lines", 11460, 1170, 2400, 450, "=V091_UI_Apply()"
    V091_AddButton oldName, "cmdFine", "Open Fine Tune", 14040, 1170, 2100, 450, "=V091_UI_OpenFineTune()"
    V091_AddButton oldName, "cmdPortfolio", "v0.9.1 Portfolio", 16320, 1170, 2200, 450, "=V091_UI_OpenPortfolio()"

    V091_AddLabel oldName, "lblJobID", "Job ID", 360, 1860, 1000, 300
    Set ctl = CreateControl(oldName, acTextBox, acDetail, , "JobID", 1560, 1800, 1800, 360)
    ctl.Name = "txtJobID": ctl.Locked = True

    V091_AddLabel oldName, "lblBCode", "Building code", 3600, 1860, 1500, 300
    V091_AddText oldName, "txtBuildingCode", "BuildingCode", 5100, 1800, 1500, ""
    V091_AddLabel oldName, "lblBName", "Building name", 6900, 1860, 1500, 300
    V091_AddText oldName, "txtBuildingName", "BuildingName", 8400, 1800, 4800, ""

    V091_AddLabel oldName, "lblClass", "Class / type", 360, 2460, 1500, 300
    Set ctl = CreateControl(oldName, acComboBox, acDetail, , "FacilityClass", 1860, 2400, 3400, 360)
    ctl.Name = "cboClass"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "B1;B1 - Lab/Rad (+50%);C1;C1 - Industrial/Rad (0%);D1;D1 - Non-Indust/Rad (-50%);D2;D2 - Non-Indust/Non-Rad (-50%)"
    ctl.ColumnCount = 2
    ctl.BoundColumn = 1
    ctl.ColumnWidths = "0;3400"
    ctl.LimitToList = True
    ctl.AfterUpdate = "=V091_UI_ClassChanged()"

    V091_AddLabel oldName, "lblType", "Type from class", 5520, 2460, 1500, 300
    Set ctl = CreateControl(oldName, acTextBox, acDetail, , "FacilityType", 7080, 2400, 2400, 360)
    ctl.Name = "txtFacilityType"
    ctl.Locked = True
    ctl.BackStyle = 0

    V091_AddLabel oldName, "lblBasis", "Estimate basis", 9780, 2460, 1600, 300
    Set ctl = CreateControl(oldName, acComboBox, acDetail, , "EstimateBasis", 11400, 2400, 2700, 360)
    ctl.Name = "cboEstimateBasis": ctl.RowSourceType = "Value List": ctl.RowSource = "Building / Structure D&D;Soil / Outdoor Remediation": ctl.LimitToList = True

    V091_AddLabel oldName, "lblArea", "Total area m2", 14400, 2460, 1500, 300
    V091_AddText oldName, "txtArea", "TotalAreaM2", 15900, 2400, 1200, "0.00"
    V091_AddLabel oldName, "lblFoot", "Footprint m2", 17280, 2460, 1500, 300
    V091_AddText oldName, "txtFoot", "FootprintAreaM2", 18780, 2400, 1200, "0.00"

    V091_AddLabel oldName, "lblDur", "Total days", 360, 3060, 1200, 300
    V091_AddText oldName, "txtDuration", "ProjectDurationDays", 1560, 3000, 1000, "0"
    V091_AddLabel oldName, "lblWork", "Work days", 2820, 3060, 1200, 300
    V091_AddText oldName, "txtWork", "WorkDays", 4020, 3000, 1000, "0"
    V091_AddLabel oldName, "lblCrew", "Crew size", 5280, 3060, 1200, 300
    V091_AddText oldName, "txtCrew", "CrewSize", 6480, 3000, 1000, "0"
    V091_AddLabel oldName, "lblAdj", "Removal adj %", 7740, 3060, 1500, 300
    V091_AddText oldName, "txtAdj", "RemovalAdjustmentPct", 9300, 3000, 1000, "0.00"
    V091_AddLabel oldName, "lblCrewFactorNote", "Crew factor = crew size / 12 is applied automatically to Summary 2026 activity rates", 10620, 3060, 3600, 300
    V091_AddCheck oldName, "chkRad", "Radiological", "IsRadiological", 14040, 3000, 1600

    V091_AddLabel oldName, "lblScope", "Scope / quantity inputs", 360, 3540, 3000, 300
    Forms(oldName).Controls("lblScope").FontBold = True
    V091_AddLabel oldName, "lblPctClean", "% clean", 360, 3900, 1200, 300
    V091_AddText oldName, "txtPctClean", "PercentClean", 1560, 3840, 1000, "0.00"
    V091_AddLabel oldName, "lblPctCont", "% contaminated", 2820, 3900, 1700, 300
    V091_AddText oldName, "txtPctCont", "PercentContaminated", 4560, 3840, 1000, "0.00"
    V091_AddLabel oldName, "lblRemovalDepth", "Removal depth m", 5880, 3900, 1700, 300
    V091_AddText oldName, "txtRemovalDepth", "RemovalDepthM", 7620, 3840, 1100, "0.00"
    V091_AddLabel oldName, "lblBack", "Backfill depth m", 9060, 3900, 1700, 300
    V091_AddText oldName, "txtBack", "BackfillDepthM", 10800, 3840, 1100, "0.00"
    V091_AddLabel oldName, "lblPipe", "Asbestos pipe m", 12240, 3900, 1700, 300
    V091_AddText oldName, "txtPipe", "AsbestosPipeLengthM", 13980, 3840, 1100, "0.00"
    V091_AddLabel oldName, "lblTile", "Asbestos tile m2", 15360, 3900, 1700, 300
    V091_AddText oldName, "txtTile", "AsbestosTileAreaM2", 17100, 3840, 1100, "0.00"
    V091_AddCheck oldName, "chkHotCell", "Include hot cell", "IncludeHotCell", 18420, 3840, 1900
    V091_AddLabel oldName, "lblHotArea", "Hot-cell area m2", 360, 4260, 1700, 300
    V091_AddText oldName, "txtHotArea", "HotCellAreaM2", 2100, 4200, 1100, "0.00"

    V091_AddLabel oldName, "lblEffort", "Excel judgement / effort controls", 360, 4440, 4200, 300
    Forms(oldName).Controls("lblEffort").FontBold = True
    V091_AddLabel oldName, "lblPMNum", "Portfolio manager number", 360, 4800, 2300, 300
    V091_AddText oldName, "txtPMNum", "PortfolioManagerNumber", 2700, 4740, 900, "0"
    V091_AddLabel oldName, "lblPMUse", "Portfolio use", 3900, 4800, 1400, 300
    V091_AddText oldName, "txtPMUse", "PortfolioManagerUseFactor", 5400, 4740, 900, "0.00"
    V091_AddLabel oldName, "lblSeniorNum", "Senior PM number", 6660, 4800, 1900, 300
    V091_AddText oldName, "txtSeniorNum", "SeniorPMNumber", 8580, 4740, 900, "0"
    V091_AddLabel oldName, "lblSeniorUse", "Senior PM use", 9780, 4800, 1500, 300
    V091_AddText oldName, "txtSeniorUse", "SeniorPMUseFactor", 11340, 4740, 900, "0.00"
    V091_AddLabel oldName, "lblWasteNum", "PM/Waste number", 12600, 4800, 1900, 300
    V091_AddText oldName, "txtWasteNum", "ProjectManagerNumber", 14520, 4740, 900, "0"
    V091_AddLabel oldName, "lblWasteUse", "PM/Waste use", 15720, 4800, 1500, 300
    V091_AddText oldName, "txtWasteUse", "ProjectManagerUseFactor", 17280, 4740, 900, "0.00"

    V091_AddLabel oldName, "lblProc", "Procedure hrs", 360, 5280, 1400, 300
    V091_AddText oldName, "txtProc", "ProcedureHours", 1800, 5220, 900, "0"
    V091_AddLabel oldName, "lblQA", "QA hrs", 2940, 5280, 900, 300
    V091_AddText oldName, "txtQA", "QASafetyHours", 3900, 5220, 900, "0"
    V091_AddCheck oldName, "chkTraining", "Include training", "IncludeTraining", 5100, 5220, 1800
    V091_AddCheck oldName, "chkConsumables", "Include consumables", "IncludeConsumables", 7080, 5220, 2200
    V091_AddCheck oldName, "chkChar", "Include detailed characterisation", "IncludeDetailedCharacterization", 9480, 5220, 3100

    V091_AddCheck oldName, "chkSitePrep", "Include site prep", "IncludeSitePrep", 360, 5760, 2000
    V091_AddLabel oldName, "lblSiteHours", "Site prep hours/task", 2640, 5820, 1900, 300
    V091_AddText oldName, "txtSiteHours", "SitePrepHoursPerTask", 4560, 5760, 900, "0"
    V091_AddLabel oldName, "lblSiteTasks", "Site prep task count", 5700, 5820, 1900, 300
    V091_AddText oldName, "txtSiteTasks", "SitePrepTaskCount", 7680, 5760, 900, "0"

    V091_AddCheck oldName, "chkSiteSurvey", "Initial site survey", "SitePrepInitialSurvey", 360, 6240, 2300
    V091_AddCheck oldName, "chkSiteBoundaries", "Setup boundaries / HEPA / seal openings", "SitePrepBoundariesHepa", 2940, 6240, 4100
    V091_AddCheck oldName, "chkSiteStaging", "Establish staging area", "SitePrepStagingArea", 7380, 6240, 2600
    V091_AddCheck oldName, "chkSiteRadSeg", "Rad/non-rad segregation", "SitePrepRadSegregation", 10320, 6240, 3000
    V091_AddCheck oldName, "chkSiteElec", "Electrical isolation", "SitePrepElectricalIsolation", 13680, 6240, 2400
    V091_AddCheck oldName, "chkSitePipe", "Piping isolation", "SitePrepPipingIsolation", 16440, 6240, 2300

    V091_AddLabel oldName, "lblCharSpec", "Characterisation specialist count", 360, 6780, 2800, 300
    V091_AddText oldName, "txtCharSpec", "CharacterizationSpecialistCount", 3240, 6720, 900, "0"
    V091_AddLabel oldName, "lblCharPM", "Characterisation PM count", 4440, 6780, 2300, 300
    V091_AddText oldName, "txtCharPM", "CharacterizationPMCount", 6840, 6720, 900, "0"
    V091_AddLabel oldName, "lblCharHours", "Characterisation hrs/person", 8040, 6780, 2600, 300
    V091_AddText oldName, "txtCharHours", "CharacterizationHoursPerPerson", 10740, 6720, 900, "0"

    V091_AddLabel oldName, "lblMonths", "Consumable months", 360, 7320, 1900, 300
    V091_AddText oldName, "txtMonths", "ConsumableMonths", 2280, 7260, 1100, "0.00"
    V091_AddLabel oldName, "lblDos", "Dosimeter years", 3720, 7320, 1700, 300
    V091_AddText oldName, "txtDos", "DosimeterYears", 5460, 7260, 1100, "0.00"
    V091_AddLabel oldName, "lblBio", "Bioassay years", 6900, 7320, 1700, 300
    V091_AddText oldName, "txtBio", "BioassayYears", 8640, 7260, 1100, "0.00"

    V091_AddLabel oldName, "lblGenTotal", "v0.9.1 generated subtotal", 360, 7920, 2800, 300
    Set ctl = CreateControl(oldName, acTextBox, acDetail, , "=IIf(Nz([txtJobID],"""")="""",0,Nz(DLookUp(""V091GeneratedSubtotalAUD"",""qryV091Totals"",""JobID='"" & [txtJobID] & ""'""),0))", 3300, 7860, 2300, 360)
    ctl.Name = "txtV091Total": ctl.Format = "Currency": ctl.Locked = True: ctl.FontBold = True
    V091_AddLabel oldName, "lblFine", "Fine tune grand total", 5940, 7920, 2400, 300
    Set ctl = CreateControl(oldName, acTextBox, acDetail, , "=IIf(Nz([txtJobID],"""")="""",0,Nz(DLookUp(""GrandTotalAUD"",""qryGrandTotals"",""JobID='"" & [txtJobID] & ""'""),0))", 8460, 7860, 2300, 360)
    ctl.Name = "txtFineTotal": ctl.Format = "Currency": ctl.Locked = True: ctl.FontBold = True

    Set ctl = CreateControl(oldName, acSubform, acDetail, , , 360, 8460, 21500, 7200)
    ctl.Name = "subV091GeneratedLines": ctl.SourceObject = "Form.frmV091GeneratedLinesSubform": ctl.LinkMasterFields = "JobID": ctl.LinkChildFields = "JobID"

    DoCmd.Save acForm, oldName
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename "frmV091GenericEstimate", acForm, oldName
End Sub

Private Sub V091_CreatePortfolioForm()
    Dim frm As Form, oldName As String
    Set frm = CreateForm
    oldName = frm.Name
    With frm
        .RecordSource = "qryV091PortfolioOverview"
        .Caption = "v0.9.1 Portfolio Overview"
        .DefaultView = 2
        .ViewsAllowed = 2
        .AllowEdits = False
        .AllowAdditions = False
        .AllowDeletions = False
        .NavigationButtons = True
        .RecordSelectors = True
    End With
    V091_AddDatasheetText oldName, "JobID", "JobID", 0, 0, 1200, True
    V091_AddDatasheetText oldName, "JobName", "JobName", 1200, 0, 3000, True
    V091_AddDatasheetText oldName, "BCode", "BuildingCode", 4200, 0, 1100, True
    V091_AddDatasheetText oldName, "Class", "FacilityClass", 5300, 0, 700, True
    V091_AddDatasheetText oldName, "Type", "FacilityType", 6000, 0, 1700, True
    V091_AddDatasheetText oldName, "Basis", "EstimateBasis", 7700, 0, 2200, True
    V091_AddDatasheetText oldName, "Area", "TotalAreaM2", 9900, 0, 1100, True, "0.00"
    V091_AddDatasheetText oldName, "Days", "ProjectDurationDays", 11000, 0, 900, True, "0"
    V091_AddDatasheetText oldName, "Crew", "CrewSize", 11900, 0, 900, True, "0"
    V091_AddDatasheetText oldName, "Adj", "RemovalAdjustmentPct", 12800, 0, 900, True, "0.00"
    V091_AddDatasheetText oldName, "V091Total", "V091GeneratedSubtotalAUD", 13700, 0, 1700, True, "Currency"
    V091_AddDatasheetText oldName, "FineSub", "FineTuneSubtotalAUD", 15400, 0, 1700, True, "Currency"
    V091_AddDatasheetText oldName, "FineGrand", "FineTuneGrandTotalAUD", 17100, 0, 1700, True, "Currency"
    V091_AddDatasheetText oldName, "Delta", "DeltaAUD", 18800, 0, 1700, True, "Currency"
    V091_AddDatasheetText oldName, "Generated", "LastGeneratedAt", 20500, 0, 1700, True, "dd-mmm-yyyy"
    DoCmd.Save acForm, oldName
    DoCmd.Close acForm, oldName, acSaveYes
    DoCmd.Rename "frmV091PortfolioOverview", acForm, oldName
End Sub

' ============================================================
' DATA HELPERS
' ============================================================

Private Sub V091_CreateInputIfMissing(ByVal jobID As String)
    If DCount("*", "tblV091BuildingInputs", "JobID=" & V091_Q(jobID)) > 0 Then Exit Sub
    Dim jobName As String
    jobName = Nz(DLookup("JobName", "tblJobs", "JobID=" & V091_Q(jobID)), jobID)

    ' v0.9.1 UI/state cleanup: only create identity fields and blank/default-off booleans.
    ' Estimating assumptions remain blank until the user enters them.
    CurrentDb.Execute "INSERT INTO tblV091BuildingInputs (JobID, BuildingCode, BuildingName, ScaleRemovalByCrew, IncludeTraining, IncludeSitePrep, IncludeDetailedCharacterization, IncludeConsumables, IsRadiological, IncludeHotCell, UpdatedAt) VALUES (" & _
        V091_Q(jobID) & ", " & V091_Q(jobID) & ", " & V091_Q(jobName) & ", False, False, False, False, False, False, False, Now());", dbFailOnError
End Sub

Private Function V091_InputDbl(ByVal jobID As String, ByVal fieldName As String, ByVal defaultValue As Double) As Double
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup(fieldName, "tblV091BuildingInputs", "JobID=" & V091_Q(jobID))
    If IsNull(v) Or v = "" Then V091_InputDbl = defaultValue Else V091_InputDbl = CDbl(v)
    Exit Function
Fallback:
    V091_InputDbl = defaultValue
End Function

Private Function V091_InputText(ByVal jobID As String, ByVal fieldName As String, ByVal defaultValue As String) As String
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup(fieldName, "tblV091BuildingInputs", "JobID=" & V091_Q(jobID))
    If IsNull(v) Or v = "" Then V091_InputText = defaultValue Else V091_InputText = CStr(v)
    Exit Function
Fallback:
    V091_InputText = defaultValue
End Function

Private Function V091_InputBool(ByVal jobID As String, ByVal fieldName As String, ByVal defaultValue As Boolean) As Boolean
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup(fieldName, "tblV091BuildingInputs", "JobID=" & V091_Q(jobID))
    If IsNull(v) Or v = "" Then V091_InputBool = defaultValue Else V091_InputBool = CBool(v)
    Exit Function
Fallback:
    V091_InputBool = defaultValue
End Function

Private Function V091_SettingDbl(ByVal settingName As String, ByVal defaultValue As Double) As Double
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup("SettingValueNumber", "tblV091Settings", "SettingName=" & V091_Q(settingName))
    If IsNull(v) Or v = "" Then V091_SettingDbl = defaultValue Else V091_SettingDbl = CDbl(v)
    Exit Function
Fallback:
    V091_SettingDbl = defaultValue
End Function

Private Function V091_BaseRate(ByVal categoryName As String, ByVal itemID As Long) As Double
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup("BaseUnitRateUSD2009", "tblCostLibrary", "CategoryName=" & V091_Q(categoryName) & " AND ItemID=" & itemID)
    If IsNull(v) Or v = "" Then V091_BaseRate = 0 Else V091_BaseRate = CDbl(v)
    Exit Function
Fallback:
    V091_BaseRate = 0
End Function

Private Function V091_LabourRate(ByVal labourItemID As Long) As Double
    V091_LabourRate = V091_BaseRate("Labour", labourItemID)
End Function

Private Function V091_V062Multiplier() As Double
    V091_V062Multiplier = V091_GetSettingDblFromV062("InflationFactor", 1) * V091_GetSettingDblFromV062("UsdToAudRate", 1) * (1 + V091_GetSettingDblFromV062("OverheadPct", 0) / 100)
    If V091_V062Multiplier = 0 Then V091_V062Multiplier = 1
End Function

Private Function V091_GetSettingDblFromV062(ByVal settingName As String, ByVal defaultValue As Double) As Double
    On Error GoTo Fallback
    Dim v As Variant
    v = DLookup("SettingValueNumber", "tblSettings", "SettingName=" & V091_Q(settingName))
    If IsNull(v) Or v = "" Then V091_GetSettingDblFromV062 = defaultValue Else V091_GetSettingDblFromV062 = CDbl(v)
    Exit Function
Fallback:
    V091_GetSettingDblFromV062 = defaultValue
End Function

Private Sub V091_MoveFormToJob(ByVal formName As String, ByVal jobID As String)
    On Error GoTo CleanFail
    Dim rs As DAO.Recordset
    Set rs = Forms(formName).RecordsetClone
    rs.FindFirst "JobID=" & V091_Q(jobID)
    If Not rs.NoMatch Then Forms(formName).Bookmark = rs.Bookmark
    rs.Close
CleanFail:
End Sub

' ============================================================
' GENERIC ACCESS HELPERS
' ============================================================

Private Sub V091_CreateTableIfMissing(ByVal tableName As String, ByVal ddl As String)
    If Not V091_TableExists(tableName) Then CurrentDb.Execute ddl, dbFailOnError
End Sub

Private Function V091_TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then V091_TableExists = True: Exit Function
    Next tdf
    V091_TableExists = False
End Function

Private Function V091_QueryExists(ByVal queryName As String) As Boolean
    On Error GoTo NotFound
    Dim qdf As DAO.QueryDef
    Set qdf = CurrentDb.QueryDefs(queryName)
    V091_QueryExists = True
    Exit Function
NotFound:
    V091_QueryExists = False
End Function

Private Sub V091_AddFieldIfMissing(ByVal tableName As String, ByVal fieldName As String, ByVal fieldTypeSql As String)
    If V091_FieldExists(tableName, fieldName) Then Exit Sub
    CurrentDb.Execute "ALTER TABLE " & tableName & " ADD COLUMN " & fieldName & " " & fieldTypeSql & ";", dbFailOnError
End Sub

Private Function V091_FieldExists(ByVal tableName As String, ByVal fieldName As String) As Boolean
    On Error GoTo MissingField
    Dim fld As DAO.Field
    Set fld = CurrentDb.TableDefs(tableName).Fields(fieldName)
    V091_FieldExists = True
    Exit Function
MissingField:
    V091_FieldExists = False
End Function

Private Sub V091_DeleteFormIfExists(ByVal formName As String)
    On Error Resume Next
    DoCmd.Close acForm, formName, acSaveNo
    DoCmd.DeleteObject acForm, formName
    On Error GoTo 0
End Sub

Private Function V091_Q(ByVal value As Variant) As String
    If IsNull(value) Then V091_Q = "Null" Else V091_Q = "'" & Replace(CStr(value), "'", "''") & "'"
End Function

Private Function V091_SqlNum(ByVal value As Double) As String
    V091_SqlNum = Replace(CStr(value), ",", ".")
End Function

Private Function V091_SqlCurrency(ByVal value As Double) As String
    V091_SqlCurrency = Replace(CStr(CCur(value)), ",", ".")
End Function

Private Sub V091_AddLabel(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long, ByVal heightVal As Long)
    Dim ctl As Control
    Set ctl = CreateControl(formName, acLabel, acDetail, , , leftPos, topPos, widthVal, heightVal)
    ctl.Name = controlName: ctl.Caption = captionText: ctl.FontSize = 9: ctl.FontBold = True
End Sub

Private Function V091_AddText(ByVal formName As String, ByVal controlName As String, ByVal sourceName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long, Optional ByVal fmt As String = "") As Control
    Dim ctl As Control
    Set ctl = CreateControl(formName, acTextBox, acDetail, , sourceName, leftPos, topPos, widthVal, 360)
    ctl.Name = controlName: ctl.ControlSource = sourceName
    If fmt <> "" Then ctl.Format = fmt
    Set V091_AddText = ctl
End Function

Private Sub V091_AddCheck(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal sourceName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long)
    Dim chk As Control, lbl As Control
    Set chk = CreateControl(formName, acCheckBox, acDetail, , sourceName, leftPos, topPos, 300, 300)
    chk.Name = controlName: chk.ControlSource = sourceName
    If captionText <> "" Then
        Set lbl = CreateControl(formName, acLabel, acDetail, , , leftPos + 360, topPos, widthVal - 360, 300)
        lbl.Name = controlName & "_Label": lbl.Caption = captionText: lbl.FontSize = 9
    End If
End Sub

Private Sub V091_AddButton(ByVal formName As String, ByVal controlName As String, ByVal captionText As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long, ByVal heightVal As Long, ByVal onClickExpression As String)
    Dim ctl As Control
    Set ctl = CreateControl(formName, acCommandButton, acDetail, , , leftPos, topPos, widthVal, heightVal)
    ctl.Name = controlName: ctl.Caption = captionText: ctl.OnClick = onClickExpression
End Sub

Private Sub V091_AddDatasheetText(ByVal formName As String, ByVal ctlName As String, ByVal sourceName As String, ByVal leftPos As Long, ByVal topPos As Long, ByVal widthVal As Long, ByVal readOnly As Boolean, Optional ByVal fmt As String = "")
    Dim ctl As Control
    Set ctl = CreateControl(formName, acTextBox, acDetail, , sourceName, leftPos, topPos, widthVal, 300)
    ctl.Name = ctlName: ctl.ControlSource = sourceName: ctl.ColumnWidth = widthVal: ctl.Locked = readOnly: ctl.Enabled = True: ctl.TabStop = Not readOnly
    If fmt <> "" Then ctl.Format = fmt
End Sub
