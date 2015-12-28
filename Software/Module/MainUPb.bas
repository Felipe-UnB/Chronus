Attribute VB_Name = "MainUPb"
Option Explicit

    'Updated 24-08-2015
    
    'NEXT TEST - reduce analyses with analyses types with more than one name
    
    'Future updates
        'Create a procedure to cell change events, which will check if any ot the cell were formatted as strikethrough and it will format all the line like this
        'Change the way the sample charts are created, letting the uyser select which charts he wants to see plotted using an userform with the options

    'Known bugs:
    'Trendline labels dont update automatically after removing some cycles.
    '.Getfolder does not list all the files in the indicated folder
    
    Public Const ProgramName = "CHRONUS"
    
    Public ShowPresentation As Range 'RAnge where the user option to show or not the Chronus presentation
    
    Public StartANDOptions_Path As String
    
    'Definition of variables in Box1_Start
    Public SampleName As Object 'Name of the samples
    Public FolderPath As Object
    Public ReductionDate As Object
    Public ReducedBy As Object
    Public ExternalStandard As Object
    Public InternalStandardCheck As Object
    Public InternalStandardName As Object
    Public Spot As Object
    Public Raster As Object
    Public Detector206MIC As Object
    Public Detector206Faraday As Object
    Public CheckData As Object
    Public BlankName As Object
    Public SamplesNames As Object 'How the samples are named in the analyses (z, zircon, spl, etc.)
    Public ExternalStandardName As Object
    Public SecondaryStandardName As Object
    Public RawNumberCycles As Object
    Public CycleDuration As Object
    Public AnalysisDate As Object
    
    'Definition of variables in Box2_UPb_Options
    Public ChoosenStandard As Object
    
    'Definition of ranges where information will be stored, in UPb Workbook
    Public SampleName_UPb As Range
    Public ReductionDate_UPb As Range
    Public ReducedBy_UPb As Range
    Public FolderPath_UPb As Range
    Public ExternalStandard_UPb As Range
        Public StandardName_UPb As Range
        Public Mineral_UPb As Range
        Public Description_UPb As Range
        Public Ratio68_UPb As Range
        Public Ratio68Error_UPb As Range
        Public Ratio75_UPb As Range
        Public Ratio75Error_UPb As Range
        Public Ratio76_UPb As Range
        Public Ratio76Error_UPb As Range
        Public Ratio82_UPb As Range
        Public Ratio82Error_UPb As Range
        Public RatioErrors12s_UPb As Range
        Public RatioErrorsAbs_UPb As Range
        Public UraniumConc_UPb As Range
        Public UraniumConcError_UPb As Range
        Public ThoriumConc_UPb As Range
        Public ThoriumConcError_UPb As Range
        Public ConcErrors12s_UPb As Range
        Public ConcErrorsAbs_UPb As Range
    Public InternalStandardCheck_UPb As Range
    Public InternalStandard_UPb As Range
    Public SpotRaster_UPb As Range
    Public Detector206_UPb As Range
    Public CheckData_UPb As Range
    Public BlankName_UPb As Range
    Public SamplesNames_UPb As Range
    Public ExternalStandardName_UPb As Range
    Public RawNumberCycles_UPb As Range
    Public CycleDuration_UPb As Range
    Public AnalysisDate_UPb As Range
    Public ErrBlank_UPb As Range
    Public ErrExtStd_UPb As Range
    Public ErrExtStdCert_UPb As Range
    Public ExtStdRepro_UPb As Range
    Public Isotope232analyzed As Range
    Public Isotope208analyzed As Range
    Public SelectedBins_UPb As Range
    
    Public NewFolderPath As String
    Public OldFolderPath As String 'Variable used to save the actual addresses of raw data. Sub UpdateFileAddresses uses it to
    'update the addresses.
    
    Public Const extension As String = "exp" 'Only files with this extension will be copied to the SamList sheet
    'There is a problem with this option because it depends on if Microsoft Windows is showing
    'file extensions or not.
    
    'Main workbook
    Public TW As Workbook 'Workbook contained inside UPb AddIn
    Public mwbk As Workbook
    
    'AddIn worksheets
    Public StandardsUPb_TW_Sh As Worksheet
    Public StartANDOptions_TW_Sh As Worksheet
    Public UnB_TW_Sh As Worksheet
    
    'Necessary worksheets
    Public StartANDOptions_Sh As Worksheet
    Public SamList_Sh As Worksheet
    Public BlkCalc_Sh As Worksheet 'Worksheet where blank signal will be stored
    Public SlpStdBlkCorr_Sh As Worksheet 'Worksheet where samples and standards ratios and uncertanties, both corrected by blank, will be stored
    Public SlpStdCorr_Sh As Worksheet 'Worksheet where samples ratios and uncertanties, corrected by standards, will be stored
    Public FinalReport_Sh As Worksheet 'Worksheet wih the final report
    Public Plot_Sh As Worksheet 'Sheet with analysis plots
    Public Plot_ShHidden As Worksheet 'Sheet with complete data from analysis for those cases when user wants to see all data again
    Public Results_Sh As Worksheet
    
    'Public CovarSheet As Worksheet 'Sheet created for pasting ranges necessary to variance-covariance matrix
    
    'Worksheets names
    Public Const StartANDOptions_Sh_Name As String = "Start-AND-Options"
    Public Const StandardsUPb_Sh_Name As String = "StandardsUPb"
    Public Const SamList_Sh_Name As String = "SamList"
    Public Const BlkCalc_Sh_Name As String = "BlkCalc" 'Worksheet where blank signal will be stored
    Public Const SlpStdBlkCorr_Sh_Name As String = "SlpStdBlkCorr" 'Worksheet where samples and standards ratios and uncertanties, both corrected by blank, will be stored
    Public Const SlpStdCorr_Sh_Name As String = "SlpStdCorr" 'Worksheet where samples ratios and uncertanties, corrected by standards, will be stored
    Public Const FinalReport_Sh_Name As String = "FinalReport"
    
    'Public CovarSheet As Worksheet 'Sheet created for pasting ranges necessary to variance-covariance matrix
    
    'Ranges of isotopes signal in raw data file stored in the AddIn worksheet
    Public TW_RawPb206Range As Range, TW_RawPb208Range As Range, TW_RawTh232Range As Range, TW_RawU238Range As Range, _
    TW_RawHg202Range As Range, TW_RawPb204Range As Range, TW_RawPb207Range As Range, TW_RawCyclesTimeRange As Range, _
    TW_AnalysisDateRange As Range
    
    'Ranges of isotopes header in raw data file stored in the AddIn worksheet
    Public TW_RawPb206HeaderRange As Range, TW_RawPb208HeaderRange As Range, TW_RawTh232HeaderRange As Range, TW_RawU238HeaderRange As Range
    Public TW_RawHg202HeaderRange As Range, TW_RawPb204HeaderRange As Range, TW_RawPb207HeaderRange As Range
    
    'Ranges of isotopes signal in raw data file
    Public RawPb206Range As Range, RawPb208Range As Range, RawTh232Range As Range, RawU238Range As Range, _
    RawHg202Range As Range, RawPb204Range As Range, RawPb207Range As Range, RawCyclesTimeRange As Range, _
    AnalysisDateRange As Range
    
    'Ranges of isotopes header in raw data file
    Public RawPb206HeaderRange As Range, RawPb208HeaderRange As Range, RawTh232HeaderRange As Range, RawU238HeaderRange As Range
    Public RawHg202HeaderRange As Range, RawPb204HeaderRange As Range, RawPb207HeaderRange As Range
    
    'Some collections to check if there is all the necessary information in Start-AND-Option sheet
    
    Public ii As Collection 'Important information - Collection of fundamental information that must be stored in Start-AND-Option sheet if the
    'workbook has already been used for calculation
    Public IIM As Collection 'Important Information Missing - Collection of missing informations from ImportantInformation above.
    Public IIName As Collection 'Names for each of the objects in IIM
    
    Public AllSamplesPath As Range 'File path of all samples, blanks and standards in SamList
    
'    'Some constants
'    Public Const RatioUranium As Single = 137.88
'    Public Const mvtoCPS As Single = 62500000
'    Public Const RatioMercury As Single = 4.35
    
    Public TW_RatioUranium_UPb As Range
    Public TW_RatioMercury_UPb As Range
    Public TW_mVtoCPS_UPb As Range
    Public TW_RatioMercury1Std As Range
    Public TW_SampleName As Range
    Public TW_BlankName As Range
    Public TW_PrimaryStandardName As Range
    
    Public RatioUranium_UPb As Range
    Public RatioMercury_UPb As Range
    Public mVtoCPS_UPb As Range
    Public RatioMercury1Std As Range
    
    'Box2_UPb_Options, page Error Propagation
    Public ErrBlank As Control, ErrExtStd As Control, ExtStdRepro As Control, ErrExtStdCert As Control
    
    'Box4_Addresses
    Public RawHg202 As Control, RawPb204 As Control, RawPb206 As Control, RawPb207 As Control, RawPb208 As Control, RawTh232 As Control, _
    RawU238 As Control, RawCyclesTime As Control, RawHg202Header As Control, RawPb204Header As Control, RawPb206Header As Control, RawPb207Header As Control, _
    RawPb208Header As Control, RawTh232Header As Control, RawU238Header As Control
    
    'Array with IDs of columns F to J in SamList. This array will be used during data reduction to know which standard and blank are related to a sample
    Public AnalysesList() As SamplesMap
    Public AnalysesList_std() As ExtStandardsMap
    
    'Arrays with cell address where the indicated sample, standard and blank names were found
    Public BlkFound() As Variant 'Array of BCO
    Public StdFound() As Variant 'Array of Std
    Public SlpFound() As Variant 'Array of samples
    Public IntStdFound() As Variant 'Array of internal standard
    
    Public Const IgnoreSymbol As String = "*"
        
    'ID, TimeFirstCycle, TimeFirstCyle, Cycles are just integers (string for the last one) related to the position of this information in AnalysesList array.
    Public RawDataFilesPaths As Integer
    Public FileName As Integer
    Public ID As Integer
    Public TimeFirstCycle As Integer
    Public Cycles As Integer
    
    Public PathsNamesIDsTimesCycles() As Variant 'Array with analyses paths (SamList column A), Sample names (SamList column B), IDs (SamList column C), Time of first Cycle (Samlist column D) and cycles choosen (Samlist column E)
    Public MapIDsRange As Range 'Range in SamList_Sh with IDs of samples in SamList map (columsn F to J)
    Public IDsRange As Range 'Range in SamList_Sh with IDs from all analyses (SamList column C)
    
    Public Type IDsTimesDifference 'UDT to store analyses IDs and difference os time between analyses
        ID As Integer
        TimeDifference As Double
    End Type
    
    'Messages to handle error while trying to open analysis files.
    Public MissingFile1 As String
    Public MissingFile2 As String
    
    Public Type SamplesMap 'UDT used to store samples map, with IDs of every analyses
        sample As Integer
        Std1 As Integer
        Std2 As Integer
        Blk1 As Integer
        Blk2 As Integer
    End Type
    
    Public Type ExtStandardsMap 'UDT used to store external standards map, with IDs of every analyses
        Std As Integer
        Blk1 As Integer
    End Type
    
    Public Type UPbStandards 'UDT to store UPb standard informations
        StandardName As String
        Mineral As String
        Description As String
        Ratio68 As Double
        Ratio68Error As Double
        Ratio75 As Double
        Ratio75Error As Double
        Ratio76 As Double
        Ratio76Error As Double
        Ratio82 As Double
        Ratio82Error As Double
        RatioErrors12s As Integer '1 if all ratio errors are 1 standard deviation or 2 if they are 2 standard deviation
        RatioErrorsAbs As Boolean 'True if all ratios errors are absolute.
        UraniumConc As Double
        UraniumConcError As Double
        ThoriumConc As Double
        ThoriumConcError As Double
        ConcErrors12s As Integer '1 if all concentration errors are 1 standard deviation or 2 if they are 2 standard deviation
        ConcErrorsAbs As Boolean 'True if all concentration errors are absolute.
    End Type
        
    'Public Const FirstLine As Integer = 3 'Row number of the first line available in SamList_Sh to paste any information (line below headers)
    'Public Const BlkCalcFirstLine As Integer = 2
    
    Public Const CalculationFirstCell As String = "BB1" 'First cell of the column where cells will be pasted inside raw data files to do somo calculations
    Public Const CalculationColumn As String = "BB" 'Column where cells will be pasted inside raw data files to do somo calculations
    
    'Constants used to set the right columns in SamList_Sh
    Public Const SamList_HeadersLine1 As Integer = 1
    Public Const SamList_HeadersLine2 As Integer = 2
    Public Const SamList_FirstLine As Integer = 3 'Row number of the first line available in SamList_Sh to paste any information (line below headers)
    Public Const SamList_FilePath As String = "A"
    Public Const SamList_FileName As String = "B"
    Public Const SamList_ID As String = "C"
    Public Const SamList_FirstCycleTime As String = "D"
    Public Const SamList_Cycles As String = "E"
    Public Const SamList_StdID As String = "F"
    Public Const SamList_BlkID As String = "G"
    Public Const SamList_SlpID As String = "H"
    Public Const SamList_Std1ID As String = "I"
    Public Const SamList_Std2ID As String = "J"
    Public Const SamList_Blk1ID As String = "K"
    Public Const SamList_Blk2ID As String = "L"
    
    'Constants used to set the right columns for isotopes signals or ratios in BlkCalc_Sh
    Public Const BlkCalc_HeaderLine As Integer = 1 'Row number of the headers
    Public Const BlkColumnID As String = "A"
    Public Const BlkSlpName As String = "B"
    Public Const BlkColumn2 As String = "C"
    Public Const BlkColumn21Std As String = "D"
    Public Const BlkColumn4 As String = "E"
    Public Const BlkColumn41Std As String = "F"
    Public Const BlkColumn6 As String = "G"
    Public Const BlkColumn61Std As String = "H"
    Public Const BlkColumn7 As String = "I"
    Public Const BlkColumn71Std As String = "J"
    Public Const BlkColumn8 As String = "K"
    Public Const BlkColumn81Std As String = "L"
    Public Const BlkColumn32 As String = "M"
    Public Const BlkColumn321Std As String = "N"
    Public Const BlkColumn38 As String = "O"
    Public Const BlkColumn381Std As String = "P"
    Public Const BlkColumn4Comm As String = "Q"
    Public Const BlkColumn4Comm1Std As String = "R"
    
    'Constants used to set the right columns (and header row) for isotopes signals or ratios in SlpStdBlkCorr
    Public Const HeaderRow As Integer = 9
    Public Const ExtStdReproRow As Integer = 1
    Public Const ColumnID As String = "A"
    Public Const ColumnSlpName As String = "B"
    
    Public Const Column75 As String = "C"
    Public Const Column751Std As String = "D"
    Public Const Column68 As String = "E"
    Public Const Column681Std As String = "F"
        Public Const ColumnWtdAvLabels As String = "F"
    Public Const Column7568Rho As String = "G"
    Public Const Column68R As String = "H"
    Public Const Column68R2 As String = "I"
    Public Const Column76 As String = "J"
    Public Const Column761Std As String = "K"
    Public Const Column2 As String = "L"
    Public Const Column21Std As String = "M"
    Public Const Column4 As String = "N"
    Public Const Column41Std As String = "O"
    Public Const Column6 As String = "P"
    Public Const Column61Std As String = "Q"
    Public Const Column7 As String = "R"
    Public Const Column71Std As String = "S"
    Public Const Column8 As String = "T"
    Public Const Column81Std As String = "U"
    Public Const Column32 As String = "V"
    Public Const Column321Std As String = "W"
    Public Const Column38 As String = "X"
    Public Const Column381Std As String = "Y"
    Public Const Column64 As String = "Z"
    Public Const Column641Std As String = "AA"
    Public Const Column74 As String = "AB"
    Public Const Column741Std As String = "AC"
    Public Const Column28 As String = "AD"
    Public Const Column281Std As String = "AE"
    
    Public Const ColumnExtStdRepro As String = "C"
    Public Const ColumnExtStd68 As String = "D"
    Public Const ColumnExtStd75 As String = "C"
    Public Const ColumnExtStd76 As String = "E"
    
    'Below are the variables that will keep the range for the standard deviation of 68, 75 and 76 from all external standards
    
    Public ExtStd68ReproHeader As Range
    Public ExtStd75ReproHeader As Range
    Public ExtStd76ReproHeader As Range
    Public ExtStd68Repro As Range
    Public ExtStd75Repro As Range
    Public ExtStd76Repro As Range
    
    Public ExtStd68MSWD As Range
    Public ExtStd75MSWD As Range
    Public ExtStd76MSWD As Range
    
    Public ExtStd68Repro1std As Range
    Public ExtStd75Repro1std As Range
    Public ExtStd76Repro1std As Range
        
    'Constants used to set the right columns (and header row) for isotopes signals or ratios in FinalReport
    Public Const FR_HeaderRow As Integer = 1
    'Public Const FR_ColumnID As String = "B"
    Public Const FR_ColumnSlpName As String = "A"
    
    Public Const FR_ColumnComments As String = "B"
    Public Const FR_Column204PbCps As String = "C"
    Public Const FR_Column206PbmV As String = "D"
    Public Const FR_ColumnUppm As String = "E"
    Public Const FR_ColumnThU As String = "F"
    Public Const FR_Column64 As String = "G"
    Public Const FR_Column641Std As String = "H"
    Public Const FR_ColumnTera238206 As String = "I"
    Public Const FR_ColumnTera2382061Std As String = "J"
    Public Const FR_ColumnTera76 As String = "K"
    Public Const FR_ColumnTera761Std As String = "L"
    Public Const FR_ColumnTera208206 As String = "M"
    Public Const FR_Column2082061Std As String = "N"
    Public Const FR_ColumnWeth75 As String = "O"
    Public Const FR_ColumnWeth751Std As String = "P"
    Public Const FR_ColumnWeth68 As String = "Q"
    Public Const FR_ColumnWeth681Std As String = "R"
    Public Const FR_ColumnWethRho As String = "S"
    Public Const FR_Column208232 As String = "T"
    Public Const FR_Column2082321Std As String = "U"
    Public Const FR_ColumnAge76 As String = "V"
    Public Const FR_ColumnAge762StdAbs As String = "W"
    Public Const FR_ColumnAge68 As String = "X"
    Public Const FR_ColumnAge682StdAbs As String = "Y"
    Public Const FR_ColumnAge75 As String = "Z"
    Public Const FR_ColumnAge752StdAbs As String = "AA"
    Public Const FR_ColumnAge208232 As String = "AB"
    Public Const FR_ColumnAge2082322StdAbs As String = "AC"
    Public Const FR_Column6876DiscPercent As String = "AD"
    Public Const FR_LastColumn As String = "AD"
    
    'Constants used to set the right columns (and header row) for isotopes signals or ratios in Plot_Sh
    Public Const Plot_HeaderRow As Integer = 1
    Public Const Plot_IDCell As String = "Q" & Plot_HeaderRow
    Public Const Plot_AnalysisName As String = "R" & Plot_HeaderRow
    
    Public Const Plot_ColumnCyclesTime As String = "A"
    Public Const Plot_Column2 As String = "B"
    Public Const Plot_Column4 As String = "C"
    Public Const Plot_Column6 As String = "D"
    Public Const Plot_Column7 As String = "E"
    Public Const Plot_Column8 As String = "F"
    Public Const Plot_Column32 As String = "G"
    Public Const Plot_Column38 As String = "H"
    Public Const Plot_Column64 As String = "I"
    Public Const Plot_Column74 As String = "J"
    Public Const Plot_Column28 As String = "K"
    Public Const Plot_Column75 As String = "L"
    Public Const Plot_Column68 As String = "M"
    Public Const Plot_Column76 As String = "N"
    
    Public Const Plot_FirstColumn As String = "A"
    Public Const Plot_LastColumn As String = "N"
    
    Public Plot_CyclesTimeRange As Range

    'Constants used to set the right columns (and header row) for isotopes signals or ratios in SlpStdCorr
    Public Const StdCorr_HeaderRow As Integer = 1
    
    Public Const StdCorr_FirstColumn As String = "A"
    Public Const StdCorr_LastColumn As String = "AE"
    
    Public Const StdCorr_ColumnID As String = "A"
    Public Const StdCorr_SlpName As String = "B"
    Public Const StdCorr_TetaFactor As String = "C"
    Public Const StdCorr_Column75 As String = "D"
    Public Const StdCorr_Column751Std As String = "E"
    Public Const StdCorr_Column68 As String = "F"
    Public Const StdCorr_Column681Std As String = "G"
    Public Const StdCorr_Column7568Rho As String = "H"
    Public Const StdCorr_Column68R As String = "I"
    Public Const StdCorr_Column68R2 As String = "J"
    Public Const StdCorr_Column76 As String = "K"
    Public Const StdCorr_Column761Std As String = "L"
    Public Const StdCorr_Column2 As String = "M"
    Public Const StdCorr_Column21Std As String = "N"
    Public Const StdCorr_Column4 As String = "O"
    Public Const StdCorr_Column41Std As String = "P"
    Public Const StdCorr_Column64 As String = "Q"
    Public Const StdCorr_Column641Std As String = "R"
    Public Const StdCorr_Column74 As String = "S"
    Public Const StdCorr_Column741Std As String = "T"
    Public Const StdCorr_ColumnF206 As String = "U"
    Public Const StdCorr_Column28 As String = "V"
    Public Const StdCorr_Column281Std As String = "W"
    
    Public Const StdCorr_Column68AgeMa As String = "X"
    Public Const StdCorr_Column68AgeMa1std As String = "Y"
    Public Const StdCorr_Column75AgeMa As String = "Z"
    Public Const StdCorr_Column75AgeMa1std As String = "AA"
    Public Const StdCorr_Column76AgeMa As String = "AB"
    Public Const StdCorr_Column76AgeMa1std As String = "AC"
    Public Const StdCorr_Column6876Conc As String = "AD"
    Public Const StdCorr_Column6875Conc As String = "AE"
    
        
'    'Constants used to set the right columns (and header row) for isotopes signals or ratios in SlpStdCorr
'    Public Const Results_HeaderRow1 As Integer = 1
'    Public Const Results_HeaderRow2 As Integer = 2
'
'    Public Const Results_SlpName As String = "A"
'    Public Const Results_ColumnID As String = "C"
'
'    Public Const Results_Column75 As String = "A"
'    Public Const Results_Column751Std As String = "B"
'    Public Const Results_Column68 As String = "C"
'    Public Const Results_Column681Std As String = "D"
'    Public Const Results_Column7568Rho As String = "E"
'    Public Const Results_Column76 As String = "F"
'    Public Const Results_Column761Std As String = "G"
'
'    Public Const Results_ColumnAge75 As String = "H"
'    Public Const Results_ColumnAge751Std As String = "I"
'    Public Const Results_ColumnAge68 As String = "J"
'    Public Const Results_ColumnAge681Std As String = "K"
'    Public Const Results_ColumnAge7568Rho As String = "L"
'    Public Const Results_ColumnAge76 As String = "M"
'    Public Const Results_ColumnAge761Std As String = "N"
'
'    Public Const Results_Column4 As String = "O"
'    Public Const Results_Column41Std As String = "P"
'    Public Const Results_Column64 As String = "Q"
'    Public Const Results_Column641Std As String = "R"
'    Public Const Results_Column74 As String = "S"
'    Public Const Results_Column741Std As String = "T"
'    Public Const Results_Column28 As String = "U"
'    Public Const Results_Column281Std As String = "V"
    
    'Constants used to set the right columns (and header row) in UPbStandard sheet
    Public Const UPbStd_CHeaderRow As Integer = 1
    Public Const UPbStd_ColumnStandardName As String = "A"
    Public Const UPbStd_ColumnMineral As String = "B"
    Public Const UPbStd_ColumnDescription As String = "C"
    Public Const UPbStd_ColumnRatio68 As String = "D"
    Public Const UPbStd_ColumnRatio68Error As String = "E"
    Public Const UPbStd_ColumnRatio75 As String = "F"
    Public Const UPbStd_ColumnRatio75Error As String = "G"
    Public Const UPbStd_ColumnRatio76 As String = "H"
    Public Const UPbStd_ColumnRatio76Error As String = "I"
    Public Const UPbStd_ColumnRatio82 As String = "J"
    Public Const UPbStd_ColumnRatio82Error As String = "K"
    Public Const UPbStd_ColumnRatioErrors12s As String = "L" '1 if all ratio errors are 1 standard deviation or 2 if they are 2 standard deviation"
    Public Const UPbStd_ColumnRatioErrorsAbs As String = "M" 'True if all ratios errors are absolute."
    Public Const UPbStd_ColumnUraniumConc As String = "N"
    Public Const UPbStd_ColumnUraniumConcError As String = "O"
    Public Const UPbStd_ColumnThoriumConc As String = "P"
    Public Const UPbStd_ColumnThoriumConcError As String = "Q"
    Public Const UPbStd_ColumnConcErrors12s As String = "R" '1 if all concentration errors are 1 standard deviation or 2 if they are 2 standard deviation
    Public Const UPbStd_ColumnConcErrorsAbs As String = "S" 'True if all concentration errors are absolute.

    Public UPbStd_StandardsNames As Range 'Range with standards names.

    'Variable used to control the opened workbook by some procedures
    Public WBSlp As Workbook
    
    Public UPbStd() As UPbStandards
    
    Public PreserveCycles As Boolean 'This variables is used to keep the cycles stored in
                                     'SamList_Cycles column.

    Public PreserveMaps As Boolean 'This variables is used to keep the SamList and StdList maps stored in
                                     'SamList.

    Public Const BlkSlp202 As Boolean = True 'This variable let the user choose if the 204Hg correction will consider all 202 of the
                                'analysis or only that without blank.
    Public Extra202 As Double 'Correction of 204 from gas using 202 both 202 from blank and from sample
    
    'Declaring decay constants
    Public Const Decay235U_yrs As Double = 0.00000000098485
    Public Const Decay238U_yrs As Double = 0.000000000155125
    Public Const Decay232Th_yrs As Double = 0.000000000049475
    
    Public Const ConfLevel As Double = 0.32
    Public Const CutOffRatio As Double = 10
    
    Public ScreenUpd As Boolean 'Used to save the previous value of application.screenupdating
                                'This is important to return the application.screenupdating to the previous state
                                
    Public ScreenAlerts As Boolean 'Used to save the previous value of application.displayalerts
                                   'This is important to return the application.displayalerts to the previous state
                                
    Public FailToOpen As Boolean 'Variable used to control if the program used to plot analysis was able to open one
    
    'The three variables below will be used to preserve the current values of the constants in Bpx2_UPb_Options.
    'This will make possible to undo any changes in these constants.
    Public Current238U235U As Double
    Public Current202Hg204Hg As Double
    Public CurrentmVCPS As Double
    
    Public CyclesBackUpArr() As String
    
    'The following variables will be used to control the Box7_FullDataReduction
    Public CheckBoxProgram0 As Boolean
    Public CheckBoxProgram1 As Boolean
    Public CheckBoxProgram2 As Boolean
    Public CheckBoxProgram3 As Boolean
    Public CheckBoxProgram4 As Boolean
    Public CheckBoxProgram5 As Boolean
    Public CheckBoxProgram6 As Boolean
    Public CheckBoxProgram7 As Boolean
    

Sub FullDataReduction()

Dim StartTime As Double
Dim EndTime As Double
Dim EndTime1 As Double
Dim EndTime2 As Double
Dim EndTime3 As Double
Dim EndTime4 As Double
Dim EndTime5 As Double
Dim EndTime6 As Double
Dim EndTime7 As Double
Dim EndTime8 As Double
Dim EndTime9 As Double
Dim EndTime10 As Double
Dim EndTime11 As Double
Dim EndTime12 As Double
Dim EndTime13 As Double
Dim EndTime14 As Double
Dim EndTime15 As Double
Dim EndTime16 As Double
Dim EndTime17 As Double

Dim DeltaTime1 As Double
Dim DeltaTime2 As Double
Dim DeltaTime3 As Double
Dim DeltaTime4 As Double
Dim DeltaTime5 As Double
Dim DeltaTime6 As Double
Dim DeltaTime7 As Double
Dim DeltaTime8 As Double
Dim DeltaTime9 As Double
Dim DeltaTime10 As Double
Dim DeltaTime11 As Double
Dim DeltaTime12 As Double
Dim DeltaTime13 As Double
Dim DeltaTime14 As Double
Dim DeltaTime15 As Double
Dim DeltaTime16 As Double
Dim DeltaTime17 As Double

Dim TotalAnalysis As Integer

On Error GoTo 0
'If you are looping over a range of cells it is often better (faster) to copy
'the range values to a variant array first and loop over that
'
'Dim dat As Variant
'Dim rng As Range
'Dim i As Long ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Set rng = [A1:A10000]
'dat = rng.Value  ' dat is now array (1 to 10000, 1 to 1)
'For i = LBound(dat, 1) To UBound(dat, 1)
'    dat(i, 1) = dat(i, 1) * 10 'or whatever operation you need to perform
'Next
'rng.Value = dat ' put new values back on sheet
            

StartTime = Timer
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    
    'MsgBox "PublicVariables"
    Call PublicVariables
        EndTime1 = Timer
        If Timer - StartTime = 0 Then
            DeltaTime1 = 0
        Else
            DeltaTime1 = Timer - StartTime
        End If
    
    'MsgBox "Load_UPbStandardsTypeList"
    Call Load_UPbStandardsTypeList
    '-----------------------------------------------
    
    Call AskToPreserveCycles
    
    'MsgBox "unprotectsheets"
    Call unprotectsheets
        EndTime2 = Timer
        If Timer - EndTime1 = Timer Then
            DeltaTime2 = 0
        Else
            DeltaTime2 = Timer - EndTime1
        End If

    'Box1_Start.Show

    'MsgBox "MacroFolderOffice2010"
    Call MacroFolderOffice2010
        EndTime3 = Timer
        If Timer - EndTime2 = Timer Then
            DeltaTime3 = 0
        Else
            DeltaTime3 = Timer - EndTime2
        End If

    mwbk.Save

    'MsgBox "CheckRawData"
    Call CheckRawData
        EndTime4 = Timer
        If Timer - EndTime3 = Timer Then
            DeltaTime4 = 0
        Else
            DeltaTime4 = Timer - EndTime3
        End If

    mwbk.Save

    'MsgBox "FirstCycleTime"
    Call FirstCycleTime
        EndTime5 = Timer
        If Timer - EndTime4 = Timer Then
            DeltaTime5 = 0
        Else
            DeltaTime5 = Timer - EndTime4
        End If

    mwbk.Save

    'MsgBox "IdentifyFileType"
    Call IdentifyFileType
        EndTime6 = Timer
        If Timer - EndTime5 = Timer Then
            DeltaTime6 = 0
        Else
            DeltaTime6 = Timer - EndTime5
        End If

    mwbk.Save

    'MsgBox "CreateStdListMap"
    Call CreateStdListMap
        EndTime7 = Timer
        If Timer - EndTime6 = Timer Then
            DeltaTime7 = 0
        Else
            DeltaTime7 = Timer - EndTime6
        End If

    mwbk.Save

    'MsgBox "CreateSamListMap"
    Call CreateSamListMap
        EndTime8 = Timer
        If Timer - EndTime7 = Timer Then
            DeltaTime8 = 0
        Else
            DeltaTime8 = Timer - EndTime7
        End If

    mwbk.Save

    'MsgBox "FormatMainSh"
    Call FormatMainSh

    'MsgBox "CalcBlank"
    Call CalcBlank
        EndTime9 = Timer
        If Timer - EndTime8 = Timer Then
            DeltaTime9 = 0
        Else
            DeltaTime9 = Timer - EndTime8
        End If

    mwbk.Save
    
    'MsgBox "CalcAllSlpStd_BlkCorr"
    Call CalcAllSlpStd_BlkCorr
        EndTime10 = Timer
        If Timer - EndTime9 = Timer Then
            DeltaTime10 = 0
        Else
            DeltaTime10 = Timer - EndTime9
        End If
    
    mwbk.Save
    
    'MsgBox "CalcAllSlp_StdCorr"
    Call CalcAllSlp_StdCorr
        EndTime11 = Timer
        If Timer - EndTime10 = Timer Then
            DeltaTime11 = 0
        Else
            DeltaTime11 = Timer - EndTime10
        End If
    
    mwbk.Save
    
    'MsgBox "FormatSamList"
    Call FormatSamList
        EndTime12 = Timer
        If Timer - EndTime11 = Timer Then
            DeltaTime12 = 0
        Else
            DeltaTime12 = Timer - EndTime11
        End If
    
    'MsgBox "FormatStartANDOptions"
    Call FormatStartANDOptions
        EndTime13 = Timer
        If Timer - EndTime12 = Timer Then
            DeltaTime13 = 0
        Else
            DeltaTime13 = Timer - EndTime12
        End If

    'MsgBox "FormatSlpStdBlkCorr"
    Call FormatSlpStdBlkCorr
        EndTime14 = Timer
        If Timer - EndTime13 = Timer Then
            DeltaTime14 = 0
        Else
            DeltaTime14 = Timer - EndTime13
        End If
    
    'MsgBox "FormatSlpStdCorr"
    Call FormatSlpStdCorr
        EndTime15 = Timer
        If Timer - EndTime14 = Timer Then
            DeltaTime15 = 0
        Else
            DeltaTime15 = Timer - EndTime13
        End If
    
    'MsgBox "FormatBlkCalc"
    Call FormatBlkCalc
        EndTime16 = Timer
        If Timer - EndTime15 = Timer Then
            DeltaTime16 = 0
        Else
            DeltaTime16 = Timer - EndTime14
        End If

    'MsgBox "protectsheets"
    Call protectsheets
            EndTime17 = Timer
            If Timer - EndTime16 = Timer Then
                DeltaTime17 = 0
            Else
                DeltaTime17 = Timer - EndTime15
            End If
    
    mwbk.Save
    
EndTime = Timer - StartTime

If AllSamplesPath Is Nothing Then
    Set AllSamplesPath = SamList_Sh.Range("A" & SamList_FirstLine, SamList_Sh.Range("A" & SamList_FirstLine).End(xlDown))
End If
TotalAnalysis = AllSamplesPath.count

MsgBox ("PublicVariables " & Round(DeltaTime1, 4) & " s " & Round(100 * DeltaTime1 / EndTime, 2) & vbNewLine & vbNewLine & _
    "unprotectsheets " & Round(DeltaTime2, 4) & " s " & Round(100 * DeltaTime2 / EndTime, 2) & vbNewLine & vbNewLine & _
    "MacroFolderOffice2010 " & Round(DeltaTime3, 4) & " s " & Round(100 * DeltaTime3 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CheckRawData " & Round(DeltaTime4, 4) & " s " & Round(100 * DeltaTime4 / EndTime, 2) & vbNewLine & vbNewLine & _
    "FirstCycleTime " & Round(DeltaTime5, 4) & " s " & Round(100 * DeltaTime5 / EndTime, 2) & vbNewLine & vbNewLine & _
    "IdentifyFileType " & Round(DeltaTime6, 4) & " s " & Round(100 * DeltaTime6 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CreateStdListMap " & Round(DeltaTime7, 4) & " s " & Round(100 * DeltaTime7 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CreateSamListMap " & Round(DeltaTime8, 4) & " s " & Round(100 * DeltaTime8 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CalcBlank " & Round(DeltaTime9, 4) & " s " & Round(100 * DeltaTime9 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CalcAllSlpStd_BlkCorr " & Round(DeltaTime10, 4) & " s " & Round(100 * DeltaTime10 / EndTime, 2) & vbNewLine & vbNewLine & _
    "CalcAllSlp_StdCorr " & Round(DeltaTime11, 4) & " s " & Round(100 * DeltaTime11 / EndTime, 2) & vbNewLine & vbNewLine & _
    "FormatSamList " & Round(DeltaTime12, 4) & " s " & Round(100 * DeltaTime12 / EndTime, 2) & vbNewLine & vbNewLine & _
    "FormatStartANDOptions " & Round(DeltaTime13, 4) & 100 * Round(DeltaTime13 / EndTime, 2) & " s " & vbNewLine & vbNewLine & _
    "FormatSlpStdBlkCorr " & Round(DeltaTime14, 4) & " s " & Round(100 * DeltaTime14 / EndTime, 2) & vbNewLine & vbNewLine & _
    "FormatSlpStdCorr " & Round(DeltaTime15, 4) & " s " & Round(100 * DeltaTime15 / EndTime, 2) & vbNewLine & vbNewLine & _
    "protectsheets " & Round(DeltaTime16, 4) & " s " & Round(100 * DeltaTime16 / EndTime, 2) & vbNewLine & vbNewLine & _
    "Total time " & Round(EndTime, 4) & " s" & vbNewLine & vbNewLine & _
    "Number of analysis " & TotalAnalysis & vbNewLine & vbNewLine & _
    "Time per analysis (Slp, Std and Blk) " & Round(EndTime / TotalAnalysis, 2) & " s")

'    Call OpenAnalysisToPlot_ByIDs(2, True)
'        Call Plot_PlotAnalysis(Worksheets("002-GJ (2)_Plot"))
'            Call LineUpMyCharts(Worksheets("002-GJ (2)_Plot"), 1)
'
'    Call OpenAnalysisToPlot_ByIDs(3, True)
'        Call Plot_PlotAnalysis(Worksheets("003-91500 (3)_Plot"))
'            Call LineUpMyCharts(Worksheets("003-91500 (3)_Plot"), 1)
'
'    Call OpenAnalysisToPlot_ByIDs(1, True)
'        Call Plot_PlotAnalysis(Worksheets("001-BCO (1)_Plot"))
'            Call LineUpMyCharts(Worksheets("001-BCO (1)_Plot"), 1)

Application.ScreenUpdating = True
Application.DisplayAlerts = True

    Call UnloadAll

End Sub

Sub CreateWorkbook()

    'For a new sample, the necessary workbook and all necessary sheets will be created.
    Dim NewSample As Integer
    Dim SaveNewWorkbook As Variant
    Dim WB As Workbook

    'Creation of the workbook used for all calculation
    NewSample = MsgBox("Would you like to start a new data reduction?", vbYesNo, "New sample")
    
    If NewSample = 6 Then
            
            Set mwbk = Workbooks.Add 'Creation of the workbook
        
            SaveNewWorkbook = False 'Related to .GetSaveAsFilneName method, which returns "FALSE" if user presses cancel button
            
            On Error Resume Next 'Important to force to code move to the next line, considering that there is some error handlers below
            
100                SaveNewWorkbook = Application.GetSaveAsFilename _
                    ("New Sample", "Excel workbook(*.xlsx),*.xlsx") 'Ask the user for a folder and a file name.
                    
                    Do While SaveNewWorkbook = False 'User hit cancel button.
                        If MsgBox("Would you like to end the program? If not, " & _
                            "please select the folder where your data is.", vbYesNo) = vbYes Then
                                mwbk.Close (False)
                                    End
                        End If
                        
                    Loop
                    
                    If Dir(SaveNewWorkbook) <> "" Then 'Mean that there is already a workbook with the name and in the folder selected by the user
                        If Not SaveNewWorkbook = False Then
                            If MsgBox("There is a workbook with the same name in this folder. Would you like to overwrite it?", vbYesNo) = vbNo Then
                                    GoTo 100
                            Else
                                For Each WB In Application.Workbooks
                                    If WB.FullName = SaveNewWorkbook Then
                                        MsgBox "There is an opened workbook with the same name, so there is the risk that it will be overwritten. Please, close it and then retry.", vbOKOnly
                                            mwbk.Close
                                                Call UnloadAll
                                                    End
                                    End If
                                Next
                                
                                ScreenAlerts = Application.DisplayAlerts
                                
                                Application.DisplayAlerts = False
                                    mwbk.SaveAs FileName:=SaveNewWorkbook, ConflictResolution:=xlLocalSessionChanges 'UPDATE
                                Application.DisplayAlerts = ScreenAlerts
                            End If
                        End If
                    End If
                                
                    Do While Dir(SaveNewWorkbook) = "" Or Err.Number <> 0 Or SaveNewWorkbook = False
                        Err.Clear
                            ScreenAlerts = Application.DisplayAlerts
                            
                                If ScreenAlerts = False Then Application.DisplayAlerts = True
                                
                                    mwbk.SaveAs FileName:=SaveNewWorkbook, ConflictResolution:=xlLocalSessionChanges 'UPDATE
                                
                                Application.DisplayAlerts = ScreenAlerts
                                
                                If Err.Number = 0 Then
                                    Exit Do
                                Else
                                    If MsgBox("Would you like to end the program?", vbYesNo) = vbYes Then
                                        mwbk.Close (False)
                                            On Error GoTo 0
                                                End
                                    Else
                                        MsgBox "Please, write a new name or change the folder.", vbOKOnly
                                            SaveNewWorkbook = Application.GetSaveAsFilename _
                                                ("New Sample", "Excel workbook(*.xlsx),*.xlsx")
                                    End If
                                End If
                                
                    Loop
                    
            On Error GoTo 0
                                                                                            
            'Start-AND-Options sheet
           
            mwbk.Sheets(1).Name = StartANDOptions_Sh_Name
                Set StartANDOptions_Sh = mwbk.Worksheets(StartANDOptions_Sh_Name)
                    
            'SamList sheet
            
            mwbk.Sheets.Add.Name = SamList_Sh_Name
                Set SamList_Sh = mwbk.Worksheets(SamList_Sh_Name)
                    SamList_Sh.Move After:=StartANDOptions_Sh
                                    
            'BlkCalc sheet

            mwbk.Sheets.Add.Name = BlkCalc_Sh_Name 'Data of the blanks
                Set BlkCalc_Sh = mwbk.Worksheets(BlkCalc_Sh_Name)
                    BlkCalc_Sh.Move After:=SamList_Sh
                                
            'SlpStdBlkCorr sheet
                            
            mwbk.Sheets.Add.Name = SlpStdBlkCorr_Sh_Name
                Set SlpStdBlkCorr_Sh = mwbk.Worksheets(SlpStdBlkCorr_Sh_Name)
                    SlpStdBlkCorr_Sh.Move After:=BlkCalc_Sh
                        
            'SlpStdCorr sheet
        
            mwbk.Sheets.Add.Name = SlpStdCorr_Sh_Name
                Set SlpStdCorr_Sh = mwbk.Worksheets(SlpStdCorr_Sh_Name)
                    SlpStdCorr_Sh.Move After:=SlpStdBlkCorr_Sh
                                                                                          
'        'Sheet addition necessary to calculate variances and covariances using VarCovar function
'            If SheetExists("VarCovar", mwbk) = False Then
'                mwbk.Sheets.Add.Name = "VarCovar" 'Addition of the sheet where values necessary to variance-covariance calculation will be pasted
'            Else
'                mwbk.Sheets("VarCovar").Cells.Clear
'            End If
        
'            Set CovarSheet = mwbk.Sheets("VarCovar")
'                CovarSheet.Visible = xlSheetHidden
             
    Else: Call UnloadAll: End
    
    End If
    
'    Box1_Start.Show

    ScreenUpd = Application.ScreenUpdating
                            
        Application.ScreenUpdating = False
        'Debug.Print Application.ScreenUpdating
        
        '---> TIP: Rather than go into debug mode and type into the immediate window, you might want to add:
        'Debug.Print Application.ScreenUpdating.
        'Immediately after the the line that turns screen updating off. That way you can be sure if it is working or not.
        
        Call FormatMainSh
            Call UnloadAll
    
        Application.ScreenUpdating = ScreenUpd
    
    Application.Goto StartANDOptions_Sh.Range("A1")
    
End Sub
Public Sub PublicVariables()
    'This sub must be called by any program.
    'Here, all necessary variable that could be used by multiple procedures must be defined. Don't forget to
    'declare it above.


    'Updated 27082015 - Now it is not necessary anymore to close any other workbook before start a new data recution.
    
    Dim Msg As Integer
    
    Set TW = ThisWorkbook
    Set mwbk = ActiveWorkbook
    
    'Worksheets from the AddIn
    Set StandardsUPb_TW_Sh = TW.Worksheets("StandardsUPb")
    Set StartANDOptions_TW_Sh = TW.Worksheets("Start-AND-Options")
    Set UnB_TW_Sh = TW.Worksheets("UnB")
    
    'The lines below try to trap na error 91, which indicates that there is not any workbook opened, so the user
    'wants to create a new workbook. These lines checks too if the opened workbook is a valid.
    'THIS IS THE FIRST ATTEMPT TO CHECK IF THIS IS A VALID CHRONUS WORKBOOK
    On Error Resume Next
    If TW.Name = mwbk.Name Then
        If Err.Number = 91 Then
            On Error GoTo 0
                Call CreateWorkbook
        Else
            MsgBox "This is not a valid workbook.", vbOKOnly
                Call CreateWorkbook
        End If
    End If

    On Error GoTo 0
    
    'BELOW IS THE SECOND ATTEMPT TO CHECK IF THIS IS A VALID CHRONUS WORKBOOK
    'Chronus tries to find its sheets, if this is not possible, this means that some
    'of them were deleted or that this is not a valid workbook.
    On Error Resume Next
        'code to identify the necessary worksheets
        Set StartANDOptions_Sh = mwbk.Sheets(StartANDOptions_Sh_Name)
        Set SamList_Sh = mwbk.Sheets(SamList_Sh_Name)
        Set BlkCalc_Sh = mwbk.Sheets(BlkCalc_Sh_Name)
        Set SlpStdBlkCorr_Sh = mwbk.Sheets(SlpStdBlkCorr_Sh_Name)
        Set SlpStdCorr_Sh = mwbk.Sheets(SlpStdCorr_Sh_Name)
        'Set FinalReport_Sh = mwbk.Sheets(FinalReport_Sh_Name)
        
        If Err.Number <> 0 Then

            MsgBox "This is not a valid workbook.", vbOKOnly
                Call CreateWorkbook
'
'            MsgBox "Please, open a valid workbook or create a new one."
'                Call UnloadAll
'                    End
        End If
    On Error GoTo 0
    
    'Code to define ranges where constants must be stored. All of them will be stored in a single file,
    'called Start-AND-Options
    Set SampleName_UPb = StartANDOptions_Sh.Range("B3")
    Set ReductionDate_UPb = StartANDOptions_Sh.Range("B4")
    Set ReducedBy_UPb = StartANDOptions_Sh.Range("B5")
    Set FolderPath_UPb = StartANDOptions_Sh.Range("B6")
    Set ExternalStandard_UPb = StartANDOptions_Sh.Range("B9")

         Set StandardName_UPb = StartANDOptions_Sh.Range("B28")
         Set Mineral_UPb = StartANDOptions_Sh.Range("B29")
         Set Description_UPb = StartANDOptions_Sh.Range("B30")
         Set Ratio68_UPb = StartANDOptions_Sh.Range("B33")
         Set Ratio68Error_UPb = StartANDOptions_Sh.Range("C33")
         Set Ratio75_UPb = StartANDOptions_Sh.Range("B34")
         Set Ratio75Error_UPb = StartANDOptions_Sh.Range("C34")
         Set Ratio76_UPb = StartANDOptions_Sh.Range("B35")
         Set Ratio76Error_UPb = StartANDOptions_Sh.Range("C35")
'         Set Ratio82_UPb =StartANDOptions_Sh.Range("
'         Set Ratio82Error_UPb =StartANDOptions_Sh.Range("
         Set RatioErrors12s_UPb = StartANDOptions_Sh.Range("D33")
         Set RatioErrorsAbs_UPb = StartANDOptions_Sh.Range("E33")
         Set UraniumConc_UPb = StartANDOptions_Sh.Range("B39")
         Set UraniumConcError_UPb = StartANDOptions_Sh.Range("C39")
         Set ThoriumConc_UPb = StartANDOptions_Sh.Range("B40")
         Set ThoriumConcError_UPb = StartANDOptions_Sh.Range("C40")
         Set ConcErrors12s_UPb = StartANDOptions_Sh.Range("D40")
         Set ConcErrorsAbs_UPb = StartANDOptions_Sh.Range("E40")

    Set InternalStandardCheck_UPb = StartANDOptions_Sh.Range("B10")
    Set InternalStandard_UPb = StartANDOptions_Sh.Range("B11")
    Set SpotRaster_UPb = StartANDOptions_Sh.Range("A14")
    Set Detector206_UPb = StartANDOptions_Sh.Range("B17")
    Set CheckData_UPb = StartANDOptions_Sh.Range("B18")
    Set BlankName_UPb = StartANDOptions_Sh.Range("D3")
    Set SamplesNames_UPb = StartANDOptions_Sh.Range("D4")
    Set ExternalStandardName_UPb = StartANDOptions_Sh.Range("D5")
    Set RawNumberCycles_UPb = StartANDOptions_Sh.Range("B55")
    Set CycleDuration_UPb = StartANDOptions_Sh.Range("B56")
    Set ErrBlank_UPb = StartANDOptions_Sh.Range("E22") 'Cell that stores the option to propagate blank uncertainties into samples
    Set ErrExtStd_UPb = StartANDOptions_Sh.Range("E23") 'Cell that stores the option to propagate primary standard analyses uncertainties into samples
    Set ErrExtStdCert_UPb = StartANDOptions_Sh.Range("E24") 'Cell that stores the option to propagate primary standard certified uncertainties into samples
    Set ExtStdRepro_UPb = StartANDOptions_Sh.Range("E25") 'Cell that stores the option to propagate primary standard uncertainties into samples based on MSWD of the analyses
    Set SelectedBins_UPb = StartANDOptions_Sh.Range("B59")
    
    'Code to set ranges in SlpStdBlkCorr_Sh
    With SlpStdBlkCorr_Sh
        Set ExtStd68ReproHeader = .Range(ColumnExtStd68 & ExtStdReproRow + 1)
        Set ExtStd75ReproHeader = .Range(ColumnExtStd75 & ExtStdReproRow + 1)
        Set ExtStd76ReproHeader = .Range(ColumnExtStd76 & ExtStdReproRow + 1)
        Set ExtStd68Repro = .Range(.Range(ColumnExtStd68 & ExtStdReproRow + 2), .Range(ColumnExtStd68 & ExtStdReproRow + 6))
        Set ExtStd75Repro = .Range(.Range(ColumnExtStd75 & ExtStdReproRow + 2), .Range(ColumnExtStd75 & ExtStdReproRow + 6))
        Set ExtStd76Repro = .Range(.Range(ColumnExtStd76 & ExtStdReproRow + 2), .Range(ColumnExtStd76 & ExtStdReproRow + 6))
        
    Set ExtStd68MSWD = ExtStd68Repro.Item(3)
    Set ExtStd75MSWD = ExtStd75Repro.Item(3)
    Set ExtStd76MSWD = ExtStd76Repro.Item(3)

    Set ExtStd68Repro1std = SlpStdBlkCorr_Sh.Range(ColumnExtStd68 & ExtStdReproRow + 3)
    Set ExtStd75Repro1std = SlpStdBlkCorr_Sh.Range(ColumnExtStd75 & ExtStdReproRow + 3)
    Set ExtStd76Repro1std = SlpStdBlkCorr_Sh.Range(ColumnExtStd76 & ExtStdReproRow + 3)

    End With
    
    'Presentation box
    Set ShowPresentation = StartANDOptions_TW_Sh.Range("B58")
        
    'Code to set the ranges for the address of each isotope signal in Start-AND-Options sheet
    Set TW_RawHg202Range = StartANDOptions_TW_Sh.Range("B45")
    Set TW_RawPb204Range = StartANDOptions_TW_Sh.Range("B46")
    Set TW_RawPb206Range = StartANDOptions_TW_Sh.Range("B47")
    Set TW_RawPb207Range = StartANDOptions_TW_Sh.Range("B48")
    Set TW_RawPb208Range = StartANDOptions_TW_Sh.Range("B49")
    Set TW_RawTh232Range = StartANDOptions_TW_Sh.Range("B50")
    Set TW_RawU238Range = StartANDOptions_TW_Sh.Range("B51")
    Set TW_RawCyclesTimeRange = StartANDOptions_TW_Sh.Range("B53")
    Set TW_AnalysisDateRange = StartANDOptions_TW_Sh.Range("B54")
    
    'Code to set the ranges for the address of each isotope header in Start-AND-Options sheet from Addin workbook
    Set TW_RawHg202HeaderRange = StartANDOptions_TW_Sh.Range("C45")
    Set TW_RawPb204HeaderRange = StartANDOptions_TW_Sh.Range("C46")
    Set TW_RawPb206HeaderRange = StartANDOptions_TW_Sh.Range("C47")
    Set TW_RawPb207HeaderRange = StartANDOptions_TW_Sh.Range("C48")
    Set TW_RawPb208HeaderRange = StartANDOptions_TW_Sh.Range("C49")
    Set TW_RawTh232HeaderRange = StartANDOptions_TW_Sh.Range("C50")
    Set TW_RawU238HeaderRange = StartANDOptions_TW_Sh.Range("C51")
    
    'Code to set the ranges for the address of each isotope signal in Start-AND-Options sheet
    Set RawHg202Range = StartANDOptions_Sh.Range("B45")
    Set RawPb204Range = StartANDOptions_Sh.Range("B46")
    Set RawPb206Range = StartANDOptions_Sh.Range("B47")
    Set RawPb207Range = StartANDOptions_Sh.Range("B48")
    Set RawPb208Range = StartANDOptions_Sh.Range("B49")
    Set RawTh232Range = StartANDOptions_Sh.Range("B50")
    Set RawU238Range = StartANDOptions_Sh.Range("B51")
    Set RawCyclesTimeRange = StartANDOptions_Sh.Range("B53")
    Set AnalysisDateRange = StartANDOptions_Sh.Range("B54")
    
    'Code to set the ranges for the address of each isotope header in Start-AND-Options sheet
    Set RawHg202HeaderRange = StartANDOptions_Sh.Range("C45")
    Set RawPb204HeaderRange = StartANDOptions_Sh.Range("C46")
    Set RawPb206HeaderRange = StartANDOptions_Sh.Range("C47")
    Set RawPb207HeaderRange = StartANDOptions_Sh.Range("C48")
    Set RawPb208HeaderRange = StartANDOptions_Sh.Range("C49")
    Set RawTh232HeaderRange = StartANDOptions_Sh.Range("C50")
    Set RawU238HeaderRange = StartANDOptions_Sh.Range("C51")
    
    'Code to set the ranges where isotopes analyzed are identified
    Set Isotope208analyzed = StartANDOptions_Sh.Range("D49")
    Set Isotope232analyzed = StartANDOptions_Sh.Range("D50")
    
    'Set range where constants from Box2_UPb_Options must be copied
    Set TW_RatioUranium_UPb = StartANDOptions_TW_Sh.Range("B22")
    Set TW_RatioMercury_UPb = StartANDOptions_TW_Sh.Range("B23")
    Set TW_mVtoCPS_UPb = StartANDOptions_TW_Sh.Range("B24")
    Set TW_RatioMercury1Std = StartANDOptions_TW_Sh.Range("C23")
    
    Set TW_BlankName = StartANDOptions_TW_Sh.Range("D3")
    Set TW_SampleName = StartANDOptions_TW_Sh.Range("D4")
    Set TW_PrimaryStandardName = StartANDOptions_TW_Sh.Range("D5")
    
    Set RatioUranium_UPb = StartANDOptions_Sh.Range("B22")
    Set RatioMercury_UPb = StartANDOptions_Sh.Range("B23")
    Set mVtoCPS_UPb = StartANDOptions_Sh.Range("B24")
    Set RatioMercury1Std = StartANDOptions_Sh.Range("C23")

    MissingFile1 = "File not found in "
    MissingFile2 = ". Please, check it and then retry."
        
End Sub

Sub Load_UPbStandardsTypeList()
    
    Dim Counter As Integer
    Dim UPbNameRng As Range
    Dim CellRow As Integer
    
    'Ranges in UPbStd sheet inside the addin
    Set UPbStd_StandardsNames = StandardsUPb_TW_Sh.Range(UPbStd_ColumnStandardName & UPbStd_CHeaderRow + 1)
        
        If IsEmpty(UPbStd_StandardsNames) = True Then
            MsgBox "You must add an standard to be able to reduce any data."
                Load Box2_UPb_Options
                    Box2_UPb_Options.Show
                        Exit Sub
        ElseIf Not IsEmpty(UPbStd_StandardsNames.Offset(1)) = True Then
            Set UPbStd_StandardsNames = StandardsUPb_TW_Sh.Range(UPbStd_StandardsNames, UPbStd_StandardsNames.End(xlDown))
        End If
    
    Counter = 1
    ReDim UPbStd(1 To 1) As UPbStandards
        For Each UPbNameRng In UPbStd_StandardsNames
            
            CellRow = UPbNameRng.Row
            
            With StandardsUPb_TW_Sh
                UPbStd(Counter).StandardName = .Range(UPbStd_ColumnStandardName & CellRow)
                UPbStd(Counter).Mineral = .Range(UPbStd_ColumnMineral & CellRow)
                UPbStd(Counter).Description = .Range(UPbStd_ColumnDescription & CellRow)
                UPbStd(Counter).Ratio68 = Val(.Range(UPbStd_ColumnRatio68 & CellRow))
                UPbStd(Counter).Ratio68Error = Val(.Range(UPbStd_ColumnRatio68Error & CellRow))
                UPbStd(Counter).Ratio75 = Val(.Range(UPbStd_ColumnRatio75 & CellRow))
                UPbStd(Counter).Ratio75Error = Val(.Range(UPbStd_ColumnRatio75Error & CellRow))
                UPbStd(Counter).Ratio76 = Val(.Range(UPbStd_ColumnRatio76 & CellRow))
                UPbStd(Counter).Ratio76Error = Val(.Range(UPbStd_ColumnRatio76Error & CellRow))
                UPbStd(Counter).Ratio82 = Val(.Range(UPbStd_ColumnRatio82 & CellRow))
                UPbStd(Counter).Ratio82Error = Val(.Range(UPbStd_ColumnRatio82Error & CellRow))
                UPbStd(Counter).RatioErrors12s = .Range(UPbStd_ColumnRatioErrors12s & CellRow)
                UPbStd(Counter).RatioErrorsAbs = .Range(UPbStd_ColumnRatioErrorsAbs & CellRow)
                UPbStd(Counter).UraniumConc = Val(.Range(UPbStd_ColumnUraniumConc & CellRow))
                UPbStd(Counter).UraniumConcError = Val(.Range(UPbStd_ColumnUraniumConcError & CellRow))
                UPbStd(Counter).ThoriumConc = Val(.Range(UPbStd_ColumnThoriumConc & CellRow))
                UPbStd(Counter).ThoriumConcError = Val(.Range(UPbStd_ColumnThoriumConcError & CellRow))
                UPbStd(Counter).ConcErrors12s = .Range(UPbStd_ColumnConcErrors12s & CellRow)
                UPbStd(Counter).ConcErrorsAbs = .Range(UPbStd_ColumnConcErrorsAbs & CellRow)
            
                Counter = Counter + 1
                
                If Not CellRow = .Range(UPbStd_ColumnStandardName & UPbStd_CHeaderRow + 1).End(xlDown).Row Then
                    ReDim Preserve UPbStd(1 To UBound(UPbStd) + 1) As UPbStandards
                End If
            
            End With
        Next

End Sub

Sub protectsheets()
    Dim Ws As Worksheet
        
        For Each Ws In mwbk.Worksheets
            Ws.Protect , False, False, False, False
        Next
End Sub

Sub unprotectsheets()
    Dim Ws As Worksheet
    
    For Each Ws In mwbk.Worksheets
        Ws.Unprotect
    Next

End Sub

Sub SelectFolder()

    'Code that let user select the folder where raw data is stored and saves the path to the Start-AND-Option workbook

    Dim strButtonCaption As String
    Dim strDialogTitle As String
    Dim SelectDialog As FileDialog
    Dim SelectionDone As Integer
    'Dim StandardFolderPath As String  - I pretend to use this variable to check if the usar has chosen some folder and not just hit the "Select a Folder" button
        
    Set SelectDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    'Captions of the SelectDialog
    strButtonCaption = "Select a Folder"
    strDialogTitle = "Folder Selection Dialog"
    
     
    With SelectDialog
        .ButtonName = strButtonCaption
        .InitialView = msoFileDialogViewDetails     'Detailed View
        .Title = strDialogTitle
        .AllowMultiSelect = False 'Let user just select only one folder
        'SelectDialog.Show displays a file dialog box and returns a Long indicating whether
        'the user pressed the Action button (-1) or the Cancel button (0).
        SelectionDone = .Show
        
        'StandardFolderPath = .InitialFileName -- My objective here was to compare the InitialFileName to the SelectesItems folder, but they are always different,
        On Error Resume Next
        FolderPath = .SelectedItems(1)
        NewFolderPath = .SelectedItems(1)
        
        If SelectionDone <> -1 Then 'The user has clicked on "Cancel" button
        
            FolderPath = ""
                           
            End
        Else
            
            FolderPath_UPb = FolderPath
            
        End If
        
          
    End With
        
End Sub

Sub CheckWorkbook()

    If mwbk Is Nothing Then
        Call PublicVariables
    End If
        
    Dim Msg As Integer

    If SheetExists(StartANDOptions_Sh_Name, mwbk) = False Or _
        SheetExists(SamList_Sh_Name, mwbk) = False Or _
        SheetExists(BlkCalc_Sh_Name, mwbk) = False Or _
        SheetExists(SlpStdBlkCorr_Sh_Name, mwbk) = False Or _
        SheetExists(SlpStdCorr_Sh_Name, mwbk) = False Then
        
            Msg = MsgBox("There are some worksheets missing. If this is the right workbook you will" & vbNewLine _
            & "have to reduce this data again. Is this the right workbook?", vbYesNo, "UPb reduction")
                
                If Msg = 7 Then
                    End
                Else: Call CreateWorkbook
                                
                End If
    End If

'    Dim a As String
'    Dim msg As Integer
'    Dim OpenedWK As String
'
'    Dim shList As Variant 'List sheets to be checked
'    Dim shListLenght As Integer 'Number of sheets in the list
'    Dim shIndex As Integer 'index of sheet in shList
'
'    on error resume Next 'This statement prevents excel to show an error if there is not any
'                         'workbook opened.
'
'    'OpenedWK must be different than "", some workbook must be opened. If its name is different than
'    'Reduction UPb, the program will question if the workbook is really UPb Reduction.
'
'    OpenedWK = Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 1))
'
'        Select Case OpenedWK
'
'            Case Is = ""
'                MsgBox "Please, open UPb Reduction workbook."
'                    End
'            Case Is <> "UPb Reduction"
'                msg = MsgBox("The name of this workbook is not UPb Reduction. Are you sure this the correct workbook?", vbYesNo, "UPb Reduction Workbook")
'                    If msg = 7 Then
'                        MsgBox "Please, open UPb Reduction workbook."
'                            End
'                    End If
'
'        End Select
'
'    'For each sheet that must be present, add it's name below
'
'    shList = Array(StartANDOptions_Sh.Name, StandardsUPb_Sh.Name, Calculation_Sh.Name, Standard1_Sh.Name, Standard2_Sh.Name, _
'    Sample_Sh.Name, Blank1_Sh.Name, Blank2_Sh.Name)
'
'    'This procedure will check the UPb data reduction workbook integrity
'    'This is done by checking the names of the sheets (yes, it's really simple)
'
'    For shIndex = 0 To UBound(shList)
'
'        a = shList(shIndex)
'        If SheetExists(a, mwbk) = False Then
'
'            MsgBox ("This workbook does not seem to contain all the necessary worksheets. #1" _
'            & vbNewLine & "Please, open an original version UPb Reduction workbook. #2")
'
'            End
'
'        End If
'
'    Next
'
'
''Some comments:
''I had a little problem because a was not defined as a string. So, when I tried to
''call it in SheetExists function a ByRef argument type mismatch occurred. Read the
''solution below:
''I suspect you haven't set up last_name properly in the caller.
''With the statement Worksheets(data_sheet).Range("C2").Value = ProcessString(last_name)
''this will only work if last_name is a string, i.e.
''Dim last_name as String appears in the caller somewhere.
''The reason for this is that VBA passes in variables by reference by default which means
''that the data types have to match exactly between caller and callee.
''Two fixes:
''1) Change your function to Public Function ProcessString(ByVal input_string As String) As String
''2) put Dim last_name As String in the caller before you use it.
''(1) works because for ByVal, a copy of input_string is taken when passing to the function which
''will coerce it into the correct data type. It also leads to better program stability since the
''function cannot modify the variable in the caller.

        
End Sub

Sub DefaultValues()

    'Updated 24-08-2015

    Dim QuestionMsgBox As Variant 'Message box displayed if something seems to be wrong in StandardsUPb sheet.
    Dim P As Variant

    'Below, the default values for each box in Box1_Start will be defined.
    
    Box1_Start.CheckBox2_CheckRawData.Value = True
    Box1_Start.TextBox8_BlankName.Value = TW_BlankName.Value
    Box1_Start.TextBox9_SamplesNames.Value = TW_SampleName.Value
    Box1_Start.TextBox10_ExternalStandardName = TW_PrimaryStandardName.Value
    Box1_Start.TextBox4.Value = Date & " (dd/mm/yyyy)" 'Inserts the date of the day when the data reduction is done.
        
    'Lines below will add the sample name based on the name of the folder where the files are
    SampleName = Dir(FolderPath_UPb, vbDirectory)
    Box1_Start.TextBox2.Value = SampleName 'Inserts a name for the sample based on the name of the folder where it is stored
    
    If FolderPath_UPb = "" Then
        FolderPath_UPb = ActiveWorkbook.path & "\" 'Original files folder path.
    Else
        Box1_Start.TextBox6.Value = FolderPath_UPb & "\"
    End If

    Call StandardsUPbComboBox

    'Default values for Error Propagation page in Box2_UPb_Options
    On Error Resume Next
    Err.Clear
    
    ErrBlank = False
    ErrExtStd = True
    ErrExtStdCert = True
    ExtStdRepro = False
            
            If Err.Number <> 0 Then
                    Set ErrBlank = Box2_UPb_Options.CheckBox3_BlankErrors
                    Set ErrExtStd = Box2_UPb_Options.CheckBox4_ExtStdErrors
                    Set ErrExtStdCert = Box2_UPb_Options.CheckBox6_CertExtStd
                    Set ExtStdRepro = Box2_UPb_Options.CheckBox5_ExtStdRepro
                    
                        ErrBlank = False
                        ErrExtStd = True
                        ErrExtStdCert = True
                        ExtStdRepro = False

            End If
    On Error GoTo 0

    'Defaul value for Check Data checkbox

    If Box1_Start.CheckBox1_InternalStandard.Value = False Then 'Checked
    
        Box1_Start.TextBox5_InternalStandardName.Enabled = False 'Not checked
        Box1_Start.TextBox5_InternalStandardName.Value = "" 'Cleans the textbox.
        Box1_Start.TextBox5_InternalStandardName.BackColor = &H8000000F 'Changes the textbox color to grey.
    
    End If
    
    Box1_Start.TextBox11_HowMany.Value = 40
    Box1_Start.TextBox12_CycleDuration = 1.042
    
    Box4_Addresses.CheckBox2 = True
    Box4_Addresses.CheckBox3 = True
    
    Box4_Addresses.RefEdit1_202 = TW_RawHg202Range.Value
    Box4_Addresses.RefEdit2_204 = TW_RawPb204Range.Value
    Box4_Addresses.RefEdit3_206 = TW_RawPb206Range.Value
    Box4_Addresses.RefEdit4_207 = TW_RawPb207Range.Value
    Box4_Addresses.RefEdit20_208 = TW_RawPb208Range.Value
    Box4_Addresses.RefEdit5_232 = TW_RawTh232Range.Value
    Box4_Addresses.RefEdit6_238 = TW_RawU238Range.Value
    Box4_Addresses.RefEdit7_202Header = TW_RawHg202HeaderRange.Value
    Box4_Addresses.RefEdit8_204Header = TW_RawPb204HeaderRange.Value
    Box4_Addresses.RefEdit9_206Header = TW_RawPb206HeaderRange.Value
    Box4_Addresses.RefEdit10_207Header = TW_RawPb207HeaderRange.Value
    Box4_Addresses.RefEdit21_208Header = TW_RawPb208HeaderRange.Value
    Box4_Addresses.RefEdit11_232Header = TW_RawTh232HeaderRange.Value
    Box4_Addresses.RefEdit12_238Header = TW_RawU238HeaderRange.Value
    Box4_Addresses.RefEdit15_CyclesTime = TW_RawCyclesTimeRange.Value
    Box4_Addresses.RefEdit22_AnalysisDate = TW_AnalysisDateRange.Value
    
    Box2_UPb_Options.CheckBox3_BlankErrors = False
    Box2_UPb_Options.CheckBox6_CertExtStd = True
    Box2_UPb_Options.CheckBox4_ExtStdErrors = True
    Box2_UPb_Options.CheckBox5_ExtStdRepro = False
    
    Box2_UPb_Options.TextBox8_RatioUranium = TW_RatioUranium_UPb.Value
    Box2_UPb_Options.TextBox9_NaturalRatioMercury = TW_RatioMercury_UPb.Value
    Box2_UPb_Options.TextBox10_MvtoCPS = TW_mVtoCPS_UPb.Value
            
End Sub

Sub OpenSetAddresses()
    'Code that let user select the folder where raw data is stored and saves the path to the Start-AND-Option workbook

    Dim strButtonCaption As String
    Dim strDialogTitle As String
    Dim SelectDialog As FileDialog
    Dim SelectionDone As Integer
    Dim RawDataFile As String
    Dim MessageBox As Variant
    Dim FilePath As String
           
    Set SelectDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    'Captions of the SelectDialog
    strButtonCaption = "Select a file"
    strDialogTitle = "File selection dialog"
    
     
1   With SelectDialog
        .ButtonName = strButtonCaption
        .InitialView = msoFileDialogViewDetails     'Detailed View
        .Title = strDialogTitle
        .AllowMultiSelect = False 'Let user just select only one folder
        SelectionDone = .Show
        
        On Error Resume Next
        RawDataFile = .SelectedItems(1)
        
        If SelectionDone <> -1 Then 'The user has clicked on "Cancel" button or didn't chose a file and clicked on ok button
        
            MessageBox = MsgBox("Do you really want to end the program?", vbYesNo)
                If MessageBox = vbYes Then
                    End
                Else
                    GoTo 1
                End If

        ElseIf Len(RawDataFile) = 0 Then
            
            MsgBox "You didn't choose a file."
                GoTo 1
        Else
        
            FolderPath_UPb.Value = FolderPath
            
        End If
        
          
    End With
    
End Sub
Sub CheckRawData()

    'This program should open every raw data file and check if:
    'all the isotopes are present (using isotopes header stored in Start-AND-Option);
    'all the isotopes have the same number of cycles;

    Dim MsgBoxAlert As Variant 'Message box for for many checks done below
    Dim C As Variant
    Dim d As Range
    Dim E As Integer
    Dim f As Range
    Dim CounterTotal As Integer
    Dim Counter As Integer
    Dim SearchStr As Long
    
    Dim AddressRawDataFile() As Variant 'Array of variables with address in Box2_UPb_Options
    Dim AddressRawDataFileHeader() As Variant 'Array of variables with address of headers in Box2_UPb_Options
    Dim Headers() As Variant
    Dim CyclesNumber As Integer 'Minimum number of cycles
    Dim OpenedWorkbook As Workbook
    Dim result1 As Variant
    Dim result2 As Variant
    Dim Result As Variant
    
    If SampleName_UPb Is Nothing Then
        Call PublicVariables
    End If
        
    If Len(SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).Value) <> 0 Or Len(SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).End(xlDown).Value) <> 0 Then
            Set AllSamplesPath = SamList_Sh.Range(SamList_FilePath & SamList_FirstLine, SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).End(xlDown))
        Else
            MsgBox "There is no samples in SamList sheet, at least not in cell " & SamList_FilePath & SamList_FirstLine & ".", vbOKOnly
                End
    End If

'    AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, RawPb208Range, RawTh232Range, RawU238Range)
'    AddressRawDataFileHeader = Array(RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, RawPb207HeaderRange, RawPb208HeaderRange, RawTh232HeaderRange, RawU238HeaderRange)
'    Headers = Array(202, 204, 206, 207, 208, 232, 238)

    'The conditional clauses below are necessary because not all isotopes must have been analyzed
        If Isotope208analyzed = True And Isotope232analyzed = True Then
            AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, _
            RawPb208Range, RawTh232Range, RawU238Range)
        ElseIf Isotope208analyzed = True And Isotope232analyzed = False Then
            AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, _
            RawPb207Range, RawPb208Range, RawU238Range)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = True Then
            AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, _
            RawPb207Range, RawTh232Range, RawU238Range)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = False Then
            AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, _
            RawPb207Range, RawU238Range)
        End If
        
        If Isotope208analyzed = True And Isotope232analyzed = True Then
            AddressRawDataFileHeader = Array(RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, _
            RawPb207HeaderRange, RawPb208HeaderRange, RawTh232HeaderRange, RawU238HeaderRange)
        ElseIf Isotope208analyzed = True And Isotope232analyzed = False Then
            AddressRawDataFileHeader = Array(RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, _
            RawPb207HeaderRange, RawPb208HeaderRange, RawU238HeaderRange)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = True Then
            AddressRawDataFileHeader = Array(RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, _
            RawPb207HeaderRange, RawTh232HeaderRange, RawU238HeaderRange)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = False Then
            AddressRawDataFileHeader = Array(RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, _
            RawPb207HeaderRange, RawU238HeaderRange)
        End If


        If Isotope208analyzed = True And Isotope232analyzed = True Then
            Headers = Array(202, 204, 206, 207, 208, 232, 238)
        ElseIf Isotope208analyzed = True And Isotope232analyzed = False Then
            Headers = Array(202, 204, 206, 207, 208, 238)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = True Then
            Headers = Array(202, 204, 206, 207, 232, 238)
        ElseIf Isotope208analyzed = False And Isotope232analyzed = False Then
            Headers = Array(202, 204, 206, 207, 238)
        End If
    
    CyclesNumber = RawNumberCycles_UPb
    
    If CheckData_UPb = True Then
        
        For Each d In AllSamplesPath
            
            On Error Resume Next
                Set OpenedWorkbook = Workbooks.Open(d)
                    If Err.Number <> 0 Then
                        '''debug.print D
                        MsgBox MissingFile1 & d & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0
            
            SearchStr = InStr(BlankName_UPb, OpenedWorkbook.Name) 'This will be used below to check if the opened data is from a blank.
                                                                                  'In this case, there´s no reason to check the signal of 206, 238, etc
            Call CentralMassCheck(OpenedWorkbook)
            
            'Check the data headers
            For E = LBound(Headers) To UBound(Headers)
                result1 = OpenedWorkbook.Sheets(1).Range(AddressRawDataFileHeader(E))
                
                'OpenedWorkbook.Activate
                
                result2 = Headers(E)
                Result = InStr(1, result1, result2, vbTextCompare)
                
                If Result = 0 Then
                    MsgBox (result2 & " is missing in " & OpenedWorkbook.Name & ". Please, check it. You may have selected the wrong range. ")
                        Application.Goto OpenedWorkbook.Worksheets(1).Range("A1")
                            Call UnloadAll
                                End
                End If
            
            Next
        
            'Check if all the ranges have the same number of cells with values (the ranges must be of the same size)
            For Each C In AddressRawDataFile
                
                E = WorksheetFunction.count(OpenedWorkbook.Worksheets(1).Range(C))
                    If E <> CyclesNumber Then
                        MsgBox ("Some cycles seem to be missing in " & OpenedWorkbook.Name & ". Please, check this file and then retry.")
                            Application.Goto OpenedWorkbook.Worksheets(1).Range("A1")
                                Call UnloadAll
                                    End
                    End If
                
            Next
            
            '206Pb signal should be > 0 in at least 50% of the cycles
            CounterTotal = 0
            Counter = 0
            
        If SearchStr > 0 Then
                For Each f In OpenedWorkbook.Worksheets(1).Range(RawPb206Range)
                    
                    If f <= 0 Then
                        Counter = Counter + 1
                    End If
                    
                    CounterTotal = CounterTotal + 1
                Next
                
                    If 100 * (Counter / CounterTotal) > 50 Then
                        MsgBox ("206Pb signal is <= 0 in " & Round(100 * (Counter / CounterTotal), 2) & "% of the cycles in " & _
                        OpenedWorkbook.Name & ". Be careful with this data.")
                    End If
                
            '238U signal should be > 0 in at least 50% of the cycles
            CounterTotal = 0
            Counter = 0
            For Each f In OpenedWorkbook.Worksheets(1).Range(RawU238Range)
                
                If f <= 0 Then
                    Counter = Counter + 1
                End If
                
                CounterTotal = CounterTotal + 1
            Next
            
                If 100 * (Counter / CounterTotal) > 50 Then
                    MsgBox ("238U signal is <= 0 in " & Round(100 * (Counter / CounterTotal), 2) & "% of the cycles in " & _
                    OpenedWorkbook.Name & ". Be careful with this data.")
                End If

         End If
         
            OpenedWorkbook.Close (True)
    
        Next
    
    End If

End Sub

Sub CentralMassCheck(WB As Workbook)
    
'    Dim WhatToLook1 As String
'    Dim WhatToLook2 As String
'    Dim result1 As Integer
'    Dim result2 As Integer
    
    Dim Test219 As Variant
    Dim Test220 As Variant
    
    With WB.Sheets(1) 'Cabeçário do arquivo da amostra
        
        With .Range("A15").EntireRow
            Set Test219 = .Find("219") 'Check if data was correctly exported
                Set Test220 = .Find("220") '
        End With
            
            If Not Test219 Is Nothing Then
                .Range(Cells(Test219.Row, Test219.Column), Cells(Test219.Row + RawNumberCycles_UPb, Test219.Column)).Delete (xlShiftToLeft)
            ElseIf Not Test220 Is Nothing Then
                .Range(Cells(Test220.Row, Test220.Column), Cells(Test220.Row + RawNumberCycles_UPb, Test220.Column)).Delete (xlShiftToLeft)
            End If
   End With
   
End Sub


Sub SetAddressess()
    'This program opens the first file path indicated by MacroFolderOffice2010 and ask for the user to select
    'the cells with all the necessary information to data reduction
    
    Dim FirstSamplePath As Range
    Dim WorkbookOpened As Workbook 'Name of the workbook opened

    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
            
    If Len(SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).Value) <> 0 Or Len(SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).End(xlDown).Value) <> 0 Then
        Set AllSamplesPath = SamList_Sh.Range(SamList_FilePath & SamList_FirstLine, SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).End(xlDown))
        Else
            Call MacroFolderOffice2010
                Set AllSamplesPath = SamList_Sh.Range(SamList_FilePath & SamList_FirstLine, SamList_Sh.Range(SamList_FilePath & SamList_FirstLine).End(xlDown))
    End If
    
    Set FirstSamplePath = AllSamplesPath.Item(1)
        
    If Len(FirstSamplePath.Value) = 0 Then
        MsgBox "There is no samples in SamList sheet. SetAddressess procedure.", vbOKOnly
            End
    End If

    On Error Resume Next
    
    ScreenUpd = Application.ScreenUpdating
        If ScreenUpd = False Then Application.ScreenUpdating = True
            
               Set WorkbookOpened = Workbooks.Open(FileName:=FirstSamplePath)
            
                If Err.Number <> 0 Then
                    MsgBox MissingFile1 & FirstSamplePath & MissingFile2 & " - " & "SetAddressess procedure"
                        Call UpdateFilesAddresses
                            Call UnloadAll
                                End
            End If
    On Error GoTo 0
                
    Box4_Addresses.Show
    
            WorkbookOpened.Close
        
    Application.ScreenUpdating = ScreenUpd
    
End Sub

Sub StandardsUPbComboBox()

    'This program will populate comboxes from Box1_Start
    'and Box2_UPb_Options will standards informations
    'stored in add-in workbook.

    Dim StandardsNamesHeader As Range 'Cell with standard names header.
    Dim Counter As Integer 'Used to add itens to External Standard ComboBox
    Dim StdRng As Range

    Box1_Start.ComboBox1_ExternalStd = ExternalStandard_UPb.Value
    Box2_UPb_Options.ComboBox1_ExternalStd = ExternalStandard_UPb.Value
    
    Box1_Start.ComboBox1_ExternalStd.Clear
    Box2_UPb_Options.ComboBox1_ExternalStd.Clear

        For Counter = 1 To UBound(UPbStd)
            Box1_Start.ComboBox1_ExternalStd.AddItem (UPbStd(Counter).StandardName)
        Next
        
        For Counter = 1 To UBound(UPbStd)
            Box2_UPb_Options.ComboBox1_ExternalStd.AddItem UPbStd(Counter).StandardName
        Next

End Sub
Sub CheckFundamentalParameters()

    'Updated 24-08-2015

    'This sub will be called everytime the user need to recalculate a grain, to verify if all the necessary information
    'are stored in the workbook
    
    Dim P As Variant 'Used to show which informations are missing in Start-AND-Options
    Dim a As Integer 'Just a counter used in "For Each" structure
    Set IIName = New Collection
    Set ii = New Collection
    Set IIM = New Collection
    
    With IIName
        .Add "What is the sample name?"
        .Add "In which folder are the raw data files?"
        .Add "Which external standard was analyzed?"
        .Add "MIC or Faraday Cup was used to analyze 206 isotope?"
        .Add "How the blanks were named?"
        .Add "How the sample analyses were named?"
        .Add "The samples were analyzed using spot or raster scheme?"
        .Add "How the standard analyses were named?"
        .Add "What is the range of 202 isotope signal in raw data file?"
        .Add "What is the range of 204 isotope signal in raw data file?"
        .Add "What is the range of 206 isotope signal in raw data file?"
        .Add "What is the range of 207 isotope signal in raw data file?"
        
        If Isotope208analyzed = True Then
            .Add "What is the range of 208 isotope signal in raw data file?"
        End If

        If Isotope232analyzed = True Then
            .Add "What is the range of 232 isotope signal in raw data file?"
        End If

        .Add "What is the range of 238 isotope signal in raw data file?"
        .Add "What is the range of 202 isotope signal header in raw data file?"
        .Add "What is the range of 204 isotope signal header in raw data file?"
        .Add "What is the range of 206 isotope signal header in raw data file?"
        .Add "What is the range of 207 isotope signal header in raw data file?"
        
        If Isotope208analyzed = True Then
            .Add "What is the range of 208 isotope signal header in raw data file?"
        End If
        
        If Isotope232analyzed = True Then
            .Add "What is the range of 232 isotope signal header in raw data file?"
        End If
        
        .Add "What is the range of 238 isotope signal header in raw data file?"
        .Add "What is the 202/204 mercury ratio? The proportion, by Rosman & Taylor (1998), is 4.35."
        .Add "What is the proportion between 238U/235U?"
        .Add "What is the conversion constant between mV and counts per second (CPS)?"
        .Add "Should the blank uncertainties be propagated into samples?"
        .Add "Should the uncertainties of the standard analyses be propagated into samples?"
        .Add "Should the certified uncertainties of the standard be propagated into samples?"
        .Add "Should the over-dispersion factor (Ibanez-Mejia et al., 2014) of the standard analyses be propagated into samples?"
    
    End With
    
    With ii
        .Add SampleName_UPb
        .Add FolderPath_UPb
        .Add ExternalStandard_UPb
        .Add Detector206_UPb
        .Add BlankName_UPb
        .Add SamplesNames_UPb
        .Add SpotRaster_UPb
        .Add ExternalStandardName_UPb
        .Add RawHg202Range
        .Add RawPb204Range
        .Add RawPb206Range
        .Add RawPb207Range

        If Isotope208analyzed = True Then
            .Add RawPb208Range
        End If

        If Isotope232analyzed = True Then
            .Add RawTh232Range
        End If
        
        .Add RawU238Range
        .Add RawHg202HeaderRange
        .Add RawPb204HeaderRange
        .Add RawPb206HeaderRange
        .Add RawPb207HeaderRange
        
        If Isotope208analyzed = True Then
            .Add RawPb208HeaderRange
        End If
        
        If Isotope232analyzed = True Then
            .Add RawTh232HeaderRange
        End If
        
        .Add RawU238HeaderRange
        .Add RatioMercury_UPb
        .Add RatioUranium_UPb
        .Add mVtoCPS_UPb
        .Add ErrBlank_UPb
        .Add ErrExtStd_UPb
        .Add ErrExtStdCert_UPb
        .Add ExtStdRepro_UPb

    End With
    
    a = 1
        For Each P In ii
            
            If P.Value = "" Then
                IIM.Add IIName.Item(a)
            End If
            a = a + 1
            
        Next

End Sub

Sub PreviousValues()

    'Updated 24-08-2015

    If SpotRaster_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    'Code to assign values from Box1_Start to the related variables in a workbook already used for data reduction
       
'    Box1_Start.TextBox10_ExternalStandardName = ExternalStandard_UPb.Value
    
    Select Case SpotRaster_UPb.Value
        Case "Spot"
            Box1_Start.OptionButton3_Spot.Value = True
        Case "Raster"
            Box1_Start.OptionButton4_Raster.Value = True
    End Select
    
    Select Case Detector206_UPb.Value
        Case "MIC"
            Box1_Start.OptionButton1_206MIC.Value = True
        Case "Faraday Cup"
            Box1_Start.OptionButton2_206Faraday.Value = True
    End Select
    
    Box1_Start.TextBox8_BlankName.Value = BlankName_UPb.Value
    Box1_Start.TextBox9_SamplesNames.Value = SamplesNames_UPb.Value
    Box1_Start.TextBox10_ExternalStandardName = ExternalStandardName_UPb.Value
    
    If InternalStandardCheck_UPb.Value = True Then
        Box1_Start.CheckBox1_InternalStandard.Value = True
        Box1_Start.TextBox5_InternalStandardName.Value = InternalStandard_UPb.Value
    End If
    
    Call StandardsUPbComboBox
    
    Box1_Start.ComboBox1_ExternalStd.Value = ExternalStandard_UPb.Value
    Box1_Start.TextBox2.Value = SampleName_UPb.Value
    Box1_Start.TextBox4.Value = ReductionDate_UPb.Value
    Box1_Start.TextBox3.Value = ReducedBy_UPb.Value
    Box1_Start.TextBox6.Value = FolderPath_UPb.Value
    
    Box1_Start.TextBox11_HowMany.Value = RawNumberCycles_UPb.Value
    Box1_Start.TextBox12_CycleDuration.Value = CycleDuration_UPb.Value
    
    Select Case CheckData_UPb.Value
        Case True
            Box1_Start.CheckBox2_CheckRawData.Value = True
        Case False
            Box1_Start.CheckBox2_CheckRawData.Value = False
    End Select
    
    'Previous addresses in Box2_UPbOptions, page addresses
    If IsEmpty(RawHg202Range) = True And _
        IsEmpty(RawPb204Range) = True And _
            IsEmpty(RawPb206Range) = True And _
                IsEmpty(RawPb207Range) = True And _
                    IsEmpty(RawPb208Range) = True And _
                        IsEmpty(RawTh232Range) = True And _
                            IsEmpty(RawU238Range) = True Then
            
            If MsgBox("Would you like to use the default cell addresses in raw data files, as well as number of cycles and duration?", vbYesNo) _
                = vbYes Then 'Aks the user if he/she wants the default values
                
                Box4_Addresses.CheckBox3.Value = True
                Box4_Addresses.CheckBox2.Value = True
            
                Box4_Addresses.RefEdit1_202.Value = TW_RawHg202Range.Value
                Box4_Addresses.RefEdit2_204.Value = TW_RawPb204Range.Value
                Box4_Addresses.RefEdit3_206.Value = TW_RawPb206Range.Value
                Box4_Addresses.RefEdit4_207.Value = TW_RawPb207Range.Value
                Box4_Addresses.RefEdit20_208.Value = TW_RawPb208Range.Value
                Box4_Addresses.RefEdit5_232.Value = TW_RawTh232Range.Value
                Box4_Addresses.RefEdit6_238.Value = TW_RawU238Range.Value
                Box4_Addresses.RefEdit7_202Header.Value = TW_RawHg202HeaderRange.Value
                Box4_Addresses.RefEdit8_204Header.Value = TW_RawPb204HeaderRange.Value
                Box4_Addresses.RefEdit9_206Header.Value = TW_RawPb206HeaderRange.Value
                Box4_Addresses.RefEdit10_207Header.Value = TW_RawPb207HeaderRange.Value
                Box4_Addresses.RefEdit21_208Header.Value = TW_RawPb208HeaderRange.Value
                Box4_Addresses.RefEdit11_232Header.Value = TW_RawTh232HeaderRange.Value
                Box4_Addresses.RefEdit12_238Header.Value = TW_RawU238HeaderRange.Value
                Box4_Addresses.RefEdit15_CyclesTime.Value = TW_RawCyclesTimeRange.Value
                Box4_Addresses.RefEdit22_AnalysisDate.Value = TW_AnalysisDateRange.Value

            End If
            
    Else 'If not all of the
        
        'AFTER SOME MODIFICATIONS IN OTHER  PROCEDURES, THE LINES BELOW RAISE AN ERROR:
        'Could not set the Value property. Type mismatch -2147352571 (80020005)
        'In order to solve this problem, a added .Value to both sides of the lines below.
        Box4_Addresses.RefEdit1_202.Value = RawHg202Range.Value
        Box4_Addresses.RefEdit2_204.Value = RawPb204Range.Value
        Box4_Addresses.RefEdit3_206.Value = RawPb206Range.Value
        Box4_Addresses.RefEdit4_207.Value = RawPb207Range.Value
        Box4_Addresses.RefEdit6_238.Value = RawU238Range.Value
        Box4_Addresses.RefEdit7_202Header.Value = RawHg202HeaderRange.Value
        Box4_Addresses.RefEdit8_204Header.Value = RawPb204HeaderRange.Value
        Box4_Addresses.RefEdit9_206Header.Value = RawPb206HeaderRange.Value
        Box4_Addresses.RefEdit10_207Header.Value = RawPb207HeaderRange.Value
        Box4_Addresses.RefEdit12_238Header.Value = RawU238HeaderRange.Value
        Box4_Addresses.RefEdit15_CyclesTime.Value = RawCyclesTimeRange.Value
        Box4_Addresses.RefEdit22_AnalysisDate.Value = AnalysisDateRange.Value

        If Isotope208analyzed = True Then
            Box4_Addresses.CheckBox3 = True
            Box4_Addresses.RefEdit20_208.Value = RawPb208Range.Value
            Box4_Addresses.RefEdit21_208Header.Value = RawPb208HeaderRange.Value
        Else
            Box4_Addresses.CheckBox3 = False
        End If
        
        If Isotope232analyzed = True Then
            Box4_Addresses.CheckBox2 = True
            Box4_Addresses.RefEdit5_232.Value = RawTh232Range.Value
            Box4_Addresses.RefEdit11_232Header.Value = RawTh232HeaderRange.Value
        Else
            Box4_Addresses.CheckBox2 = False
        End If

    End If
    
    'Constants previous values
    
    If WorksheetFunction.IsNumber(RatioMercury_UPb) = True And RatioMercury_UPb > 0 Then
        Box2_UPb_Options.TextBox9_NaturalRatioMercury.Value = RatioMercury_UPb.Value
    Else
        MsgBox "The 202Hg/204Hg constant is not a number or it is <=0. Please, check it.", vbOKOnly
    End If
        
    If WorksheetFunction.IsNumber(RatioUranium_UPb) = True And RatioUranium_UPb > 0 Then
        Box2_UPb_Options.TextBox8_RatioUranium.Value = RatioUranium_UPb.Value
    Else
        MsgBox "The 238U/235U constant is not a number or it is <=0. Please, check it.", vbOKOnly
    End If
    
    If WorksheetFunction.IsNumber(mVtoCPS_UPb) = True And mVtoCPS_UPb > 0 Then
        Box2_UPb_Options.TextBox10_MvtoCPS.Value = mVtoCPS_UPb.Value
    Else
        MsgBox "The mV to CPS constant is not a number or it is <=0. Please, check it.", vbOKOnly
    End If
    
    'Previous values for error propagation (Page Error propagation in Box2_UPb_Options
    On Error Resume Next
    
    Err.Clear
    
    ErrBlank = ErrBlank_UPb
    ErrExtStd = ErrExtStd_UPb
    ErrExtStdCert = ErrExtStdCert_UPb
    ExtStdRepro = ExtStdRepro_UPb
    
            If Err.Number <> 0 Then 'If the objects were not set yet
                    Set ErrBlank = Box2_UPb_Options.CheckBox3_BlankErrors
                    Set ErrExtStd = Box2_UPb_Options.CheckBox4_ExtStdErrors
                    Set ErrExtStdCert = Box2_UPb_Options.CheckBox6_CertExtStd
                    Set ExtStdRepro = Box2_UPb_Options.CheckBox5_ExtStdRepro
                    
                        ErrBlank = ErrBlank_UPb
                        ErrExtStd = ErrExtStd_UPb
                        ErrExtStdCert = ErrExtStdCert_UPb
                        ExtStdRepro = ExtStdRepro_UPb

            End If
    On Error GoTo 0

End Sub

Sub MacroFolderOffice2010()
' '******************************************************
'/----------------------------------------------------\
'|  Macro desenvolvida por: Felipe Valença de Oliveira|
'|  Laboratório de Geocronologia - UnB                |
'|  Primeira versão (v1): Abril - 2012                |
'|  Atualizada (v2): Junho - 2014                     |
'\----------------------------------------------------/
'******************************************************
        
'This sub makes a list of all the files in the indicated folder, looking for those necessary
'to data reduction.

    Dim FSO As Object 'Scripting.FileSystemObject
    Dim fld As Object 'Scripting.Folder
    Dim fl As Object 'Scripting.File
    Dim n As Long 'Necessary to indicate where the file path and the file name with extension will be copied
    Dim a As Long 'Number of files with the specified extension found
    Dim RemoveExtension As String 'String to be removed from the name of the sample
        
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If

    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
        Set fld = FSO.GetFolder(FolderPath_UPb)
        
        If Err.Number <> 0 Then
            MsgBox "Please, check the address of the folder where your data is."
                Call UpdateFilesAddresses
                    Call UnloadAll
                        End
        End If
    On Error GoTo 0
                
    RemoveExtension = ".dynamic.exp"
              
    With SamList_Sh
        
        If PreserveCycles = False Then
            .Cells.Clear
        End If
        
        n = SamList_FirstLine
        a = 0
        For Each fl In fld.Files
            If FSO.getExtensionName(fl.path) = extension Then
                .Cells(n, SamList_FilePath) = fl.path 'Copies the file address
                    .Cells(n, SamList_FileName) = Replace(FSO.GetFileName(fl.path), RemoveExtension, "") 'Copy the file name (with extension)
                        .Cells(n, SamList_ID) = n - 2 'Index for the sample
                
                        n = n + 1
                        a = a + 1
            End If
        Next fl
        
        Set AllSamplesPath = SamList_Sh.Range(SamList_FilePath & SamList_FirstLine, SamList_FilePath & n)
        
        
    End With
    
    If a = 0 Then
        MsgBox ("There is no file in the selected folder with the extension -->" & extension & "<-- ! " & _
        "Please, select the correct folder and then retry")
            End
            
    Else
        
    End If
      
End Sub

Sub FirstCycleTime()

    'This program opens every raw data file and copies the first cycle time to SamList sheet
    'Then it calls WriteCycles to copy the cycles IDs that will be used on the following calculations

    'Application.ScreenUpdating = False
    ''Application.DisplayAlerts = False

    Dim WorkbookOpened As Workbook
    Dim cell As Range

    If FolderPath_UPb Is Nothing Then 'We need some public variables, so we must be shure that they were set
        Call PublicVariables
    End If

    If AllSamplesPath Is Nothing Then 'This is the range in SamList with the file paths of all raw data files, so it must be defined.
        Set AllSamplesPath = SamList_Sh.Range("A" & SamList_FirstLine, SamList_Sh.Range("A" & SamList_FirstLine).End(xlDown))
    End If
           
    'Every analysis in AllSamplesPath range will be opened and the time of analysis will be copied to SamList_Sh
    For Each cell In AllSamplesPath
    
        If cell = "" Then
            Exit For
        End If

        On Error Resume Next
            Workbooks.Open FileName:=cell
                If Err.Number <> 0 Then
                    MsgBox MissingFile1 & cell & MissingFile2
                        Call UpdateFilesAddresses
                            Call UnloadAll
                                End
                End If
        On Error GoTo 0
    
            Set WorkbookOpened = ActiveWorkbook
            
                 SamList_Sh.Range("D" & cell.Row) = DateTimeCustomFormat(WorkbookOpened, WorkbookOpened.Worksheets(1).Range(RawCyclesTimeRange).Item(1), _
                 WorkbookOpened.Worksheets(1).Range(AnalysisDateRange), "hh:mm:ss:ms(xxx)", "Date: dd/mm/yyyy")
                                
                    Call WriteCycles(WorkbookOpened.Sheets(1).Range(RawCyclesTimeRange), cell.Row)

                    WorkbookOpened.Close (False)
    Next
        
    ''Application.DisplayAlerts = true
    'Application.ScreenUpdating = True

End Sub

Sub WriteCycles(CyclesRange As Range, PasteRow As Integer)

    'This program will write the number of each cycle selected by the user in the CyclesRange. It will use the cycles
    'time range (START-AND-OPTIONS sheet).
    'Arguments:
        'CyclesRanges - Range with the cycles (IDs or times)
        'PasteRow - Row of the analysis, the same as the ID of the analysis
        'NeedRecalculate means that the user removed some cycles and now the calculation must be done again. The program
        'will add an * to the end of the selected cycles.

    Dim cell2 As Range
    Dim a As Range 'Range of cycles time in raw data file (complete path)
    Dim B As Range 'Cell address of the cycle time from the sample being verified
    Dim C As Integer 'Correct cycle number, after removing d
    Dim d As Integer 'Integer with the difference between the row of the first cycle time in raw data file and row 1. This is
                     'necessary because the Cycles are written from 1 to n.
    Dim RangeItem As Integer
    
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If

    Set B = SamList_Sh.Range(SamList_Cycles & PasteRow)
    
    If PreserveCycles = True Then
            
           Exit Sub
    
    ElseIf PreserveCycles = False Then
    
        B.Clear
    
    End If
    
    d = CyclesRange.Row - 1
    
'    Set B = SamList_Sh.Range(SamList_Cycles & PasteRow)
'            B.Clear
    RangeItem = 1
    
    For Each cell2 In CyclesRange
        
        'MsgBox RangeItem & "item  selected"
        C = cell2.Row - d
                
        If Not IsEmpty(cell2) = True Then 'If there isn't information in a specific cell, we have to skip it.
            B = B & C & "," 'Adds the cycle number plus comma in cells from column E
        End If
        
    RangeItem = RangeItem + 1
    
    Next
    
    If Right(B.Value, 1) = "," Then
        B = Left(B.Value, Len(B.Value) - 1)
    End If
        
End Sub

Sub CreateStdListMap()

    'Based on informations from columns C and D, from SamList_Sh, this program creates a kind of map, in columns F to G in SamList_Sh, where samples IDs are written and
    'standards and blanks IDs related to each samples are written too.

    Dim a As Variant
    Dim C As Integer
    Dim ca As Integer 'ca and cb are counters used to populate After and Before arrays
    Dim cb As Integer
    Dim Before() As IDsTimesDifference 'Array with IDs and time differences between samples and blanks or external standards analysed BEFORE samples
    Dim After() As IDsTimesDifference 'Array with IDs and time differences between samples and blanks or external standards analysed AFTER samples
    Dim Blanks() As Integer 'Array with blanks IDs
    Dim Samples() As Integer 'Array with samples and internals standards IDs
    Dim ExtStandards() As Integer 'Array with external standards IDs
    Dim B As Double 'Variable used to compare times os analyses
    Dim d As Variant
    Dim FindIDObj As Object 'Variable used to find IDs in column F
    Dim Msg1 As String, msg2 As String, msg3 As String, msg4 As String
    
    Const StdIDColumn As Integer = 6 'Column in SamList_Sh where the external standards ID are copied to create the StdMap
    
    Call AskToPreserveListMaps
    
    If PreserveMaps = True Then
        Exit Sub
    End If
    
    If FolderPath_UPb Is Nothing Then 'We need some public variables, so we must be shure that they were set
        Call PublicVariables
    End If
    
    'Considering that we must know which analyses are from samples, standards and blanks, BlkFound must not be empty
    If IsArrayEmpty(BlkFound) = True Then
        Call IdentifyFileType
    End If
    
    C = SamList_HeadersLine2 'Integer used to set the correct row to paste IDs as defined below
    
    'Below, IDs from standards will be copied to column F in SamList_Sh and to Samples array
    ReDim Preserve ExtStandards(1 To UBound(StdFound) + 1) As Integer 'It's necessary to add 2 because the arrays from FindStrings function are dimensioned from 0 to n
    
    'Internal standards are treated like normal samples, so samples and internal standards IDs are copeid to column F in SamList_Sh
    For Each a In StdFound 'Every ID from samples will be placed in column F
        SamList_Sh.Cells(C + 1, Range(a).Offset(, 4).Column) = SamList_Sh.Range(a).Offset(, 1)
            ExtStandards(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
        
    'The code line below sorts the column F (samples and internals standards IDs) in ascending order
    SamList_Sh.Range("F" & SamList_FirstLine, SamList_Sh.Range("F" & SamList_FirstLine).End(xlDown)).Sort key1:=SamList_Sh.Range("F" & SamList_FirstLine, _
    SamList_Sh.Range("F" & SamList_FirstLine).End(xlDown)), order1:=xlAscending

        
    'Below, IDs from blanks will be copied to Blanks array
    ReDim Preserve Blanks(1 To UBound(BlkFound) + 1) As Integer 'It's necessary to add 1 because the arrays from FindStrings function are dimensioned from 0 to n

    C = 2
    
    For Each a In BlkFound
        Blanks(C - 1) = SamList_Sh.Range(a).Offset(, 1) 'Blanks IDs are copied to a different array (Blanks) which accepts only numbers (IDs)
        C = C + 1
    Next

    'Below, IDs from external standards will be copied to ExtStandards array
    ReDim Preserve ExtStandards(1 To UBound(StdFound) + 1) As Integer 'It's necessary to add 1 because the arrays from FindStrings function are dimensioned from 0 to n

    C = 2
                        
    Call SetPathsNamesIDsTimesCycles
    
    'For each standard in column F the program will choose which blank will be used.
    'This is done based on the time of first cycle, considered as the time of analysis.
    
    Msg1 = "Please, select from the dropdown list."
    msg2 = "Integers only!"
    msg3 = "You must choose a value available in the dropdownlist only!"
    msg4 = "Wrong value!"
    
    For Each a In ExtStandards
    
        ca = 1
        cb = 1
        
    'BLANKS
    ReDim Before(1 To 1) As IDsTimesDifference 'Array where all blanks or standards IDs and time differences analysed before samples will be stored
    ReDim After(1 To 1) As IDsTimesDifference 'Array where all blanks or standards IDs and time differences analysed after samples will be stored
    
        For Each d In Blanks
            B = PathsNamesIDsTimesCycles(4, a) - PathsNamesIDsTimesCycles(4, d)
                If B = 0 Then
                    MsgBox "Analysis with ID " & PathsNamesIDsTimesCycles(3, d) & " time of analyses is exactly the same of " & PathsNamesIDsTimesCycles(3, a) & _
                    ". Please, check this and then retry"
'                        Call UnloadAll
'                            End
                ElseIf B < 0 Then
                    After(ca).ID = d 'ID of the blank analyses
                    After(ca).TimeDifference = Abs(B) 'Difference of time of analyses between Blank and sample
                        ReDim Preserve After(1 To UBound(After) + 1)
                            ca = ca + 1
                Else
                    Before(cb).ID = d 'ID of the blank analyses
                    Before(cb).TimeDifference = B 'Difference of time of analyses between Blank and sample
                        ReDim Preserve Before(1 To UBound(Before) + 1)
                            cb = cb + 1
                End If
                
        Next
        
        'There must be a blank analysis before the standard and this will be checked below
        If Before(1).ID = 0 Then
            
            MsgBox "Before sample " & PathsNamesIDsTimesCycles(2, a) & " there is not any " & _
            "blank analysis. Please, check this " & _
            "missing analysis, its name might be different than the one you selected " & _
            "(" & BlankName_UPb & ").", vbOKOnly
                
                Call FormatMainSh
                    Call UnloadAll
                        SamList_Sh.Activate
                            End

        End If
        
        'Sub to sort ascending the time differences
'        Call SortMyData(After)
        Call SortMyData(Before)
        
        'The blanks or standard analyses done closer to samples is selected and so is copied to columns I to J
        With SamList_Sh.Columns(StdIDColumn)
            Set FindIDObj = .Find(a)
                
                'There are some lines below that add valitadion to SamListMap cells (columns I to J).
                
                With SamList_Sh.Range(FindIDObj.Address).Offset(, 1)
                    .Value = Before(2).ID 'Number 2 is selected because the array was sorted and the empty element (=0) went
                    'to the first position of the array
                    .Validation.Delete
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:=Join(IntegerToStringArray(Blanks), ",")
                    .Validation.InputMessage = Msg1
                    .Validation.InputTitle = msg2
                    .Validation.ErrorMessage = msg3
                    .Validation.ErrorTitle = msg4
                End With
                
        End With
        
'    'This is exactly the same code as above but is used to deal with external standards analyses
'        ca = 1
'        cb = 1
'
    Next
    
End Sub

Sub CreateSamListMap()

    'Based on informations from columns C and D, from SamList_Sh, this program creates a kind of map, in columns F to G in SamList_Sh, where samples IDs are written and
    'standards and blanks IDs related to each samples are written too.

    Dim a As Variant
    Dim C As Integer
    Dim ca As Integer 'ca and cb are counters used to populate After and Before arrays
    Dim cb As Integer
    Dim Before() As IDsTimesDifference 'Array with IDs and time differences between samples and blanks or external standards analysed BEFORE samples
    Dim After() As IDsTimesDifference 'Array with IDs and time differences between samples and blanks or external standards analysed AFTER samples
    Dim Blanks() As Integer 'Array with blanks IDs
    Dim Samples() As Integer 'Array with samples and internals standards IDs
    Dim ExtStandards() As Integer 'Array with external standards IDs
    Dim B As Double 'Variable used to compare times os analyses
    Dim d As Variant
    Dim FindIDObj As Object 'Variable used to find IDs in column F
    Dim Msg1 As String, msg2 As String, msg3 As String, msg4 As String
    Dim StdSlpColumn As String 'Column where samples IDs must be copied to in SamListMap
    Const StdSlpColumnNum As Integer = 8 'Column number where samples IDs must be copied to in SamListMap
    
    Call AskToPreserveListMaps
    
    If PreserveMaps = True Then
        Exit Sub
    End If
    
    If FolderPath_UPb Is Nothing Then 'We need some public variables, so we must be shure that they were set
        Call PublicVariables
    End If
    
    'Considering that we must know which analyses are from samples, standards and blanks, BlkFound must not be empty
    If IsArrayEmpty(BlkFound) = True Then
        Call IdentifyFileType
    End If
    
    StdSlpColumn = "H" & SamList_FirstLine 'update
    
    C = SamList_HeadersLine2 'Integer used to set the correct row to paste IDs as defined below
    
    'Below, IDs from samples and internal standards will be copied to column F in SamList_Sh and to Samples array
    ReDim Preserve Samples(1 To UBound(SlpFound) + UBound(IntStdFound) + 2) As Integer 'It's necessary to add 2 because the arrays from FindStrings function are dimensioned from 0 to n
    
    'Internal standards are treated like normal samples, so samples and internal standards IDs are copeid to column F in SamList_Sh
    For Each a In SlpFound 'Every ID from samples will be placed in column F
        SamList_Sh.Cells(C + 1, Range(a).Offset(, 6).Column) = SamList_Sh.Range(a).Offset(, 1)
            Samples(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
    
    If InternalStandardCheck_UPb = True Then
        For Each a In IntStdFound 'Every ID of internal standard will be placed in column F
            SamList_Sh.Cells(C + 1, Range(a).Offset(, 6).Column) = SamList_Sh.Range(a).Offset(, 1)
                Samples(C - 1) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
    End If
    
    'The code line below sorts the column F (samples and internals standards IDs) in ascending order
    SamList_Sh.Range(StdSlpColumn, SamList_Sh.Range(StdSlpColumn).End(xlDown)).Sort key1:=SamList_Sh.Range(StdSlpColumn, SamList_Sh.Range(StdSlpColumn).End(xlDown)), order1:=xlAscending

        
    'Below, IDs from blanks will be copied to Blanks array
    ReDim Preserve Blanks(1 To UBound(BlkFound) + 1) As Integer 'It's necessary to add 1 because the arrays from FindStrings function are dimensioned from 0 to n

    C = SamList_HeadersLine2
    
    For Each a In BlkFound
        Blanks(C - 1) = SamList_Sh.Range(a).Offset(, 1) 'Blanks IDs are copied to a different array (Blanks) which accepts only numbers (IDs)
        C = C + 1
    Next

    'Below, IDs from external standards will be copied to ExtStandards array
    ReDim Preserve ExtStandards(1 To UBound(StdFound) + 1) As Integer 'It's necessary to add 1 because the arrays from FindStrings function are dimensioned from 0 to n

    C = SamList_HeadersLine2
    
    For Each a In StdFound 'External standardsd IDs are copied to a different array (Blanks) which accepts only numbers (IDs)
        ExtStandards(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
                    
    Call SetPathsNamesIDsTimesCycles
    
    'For each sample in column F (samples and internal standards, the program will choose which blank and external standard will be used.
    'This is done based on the time of first cycle, considered as the time of analysis.
    
    Msg1 = "Please, select from the dropdown list."
    msg2 = "Integers only"
    msg3 = "You must choose a value available in the dropdownlist only!"
    msg4 = "Wrong value"
    
    For Each a In Samples
    
        ca = 1
        cb = 1
        
    'BLANKS
    ReDim Before(1 To 1) As IDsTimesDifference 'Array where all blanks or standards IDs and time differences analysed before samples will be stored
    ReDim After(1 To 1) As IDsTimesDifference 'Array where all blanks or standards IDs and time differences analysed after samples will be stored
    
        For Each d In Blanks
            B = PathsNamesIDsTimesCycles(4, a) - PathsNamesIDsTimesCycles(4, d)
                If B = 0 Then
                    MsgBox PathsNamesIDsTimesCycles(3, d) & "time of analyses is exactly the same of " & PathsNamesIDsTimesCycles(3, a) & _
                    ". Please, check this and then retry."
                ElseIf B < 0 Then
                    After(ca).ID = d 'ID of the blank analyses
                    After(ca).TimeDifference = Abs(B) 'Difference of time of analyses between Blank and sample
                        ReDim Preserve After(1 To UBound(After) + 1)
                            ca = ca + 1
                Else
                    Before(cb).ID = d 'ID of the blank analyses
                    Before(cb).TimeDifference = B 'Difference of time of analyses between Blank and sample
                        ReDim Preserve Before(1 To UBound(Before) + 1)
                            cb = cb + 1
                End If
                
        Next
        
        'There must be a blank analysis before and after the sample and this will be checked below
        If After(1).ID = 0 Then
            
            MsgBox "After sample " & PathsNamesIDsTimesCycles(2, a) & " there is not any " & _
            "blank analysis. Please, check this " & _
            "missing analysis, its name might be different than the one you selected " & _
            "(" & BlankName_UPb & ").", vbOKOnly
                
                Call FormatMainSh
                    Call UnloadAll
                        SamList_Sh.Activate
                            End
                        
        ElseIf Before(1).ID = 0 Then
            
            MsgBox "Before sample " & PathsNamesIDsTimesCycles(2, a) & " there is not any " & _
            "blank analysis. Please, check this " & _
            "missing analysis, its name might be different than the one you selected " & _
            "(" & BlankName_UPb & ").", vbOKOnly
                
                Call FormatMainSh
                    Call UnloadAll
                        SamList_Sh.Activate
                            End

        End If
        
        'Sub to sort ascending the time differences
        Call SortMyData(After)
        Call SortMyData(Before)
        
        'The blanks or standard analyses done closer to samples is selected and so is copied to columns I to J
        With SamList_Sh.Columns(StdSlpColumnNum)
            Set FindIDObj = .Find(a)
                
                'There are some lines below that add valitadion to SamListMap cells (columns I to J).
                
                With SamList_Sh.Range(FindIDObj.Address).Offset(, 3)
                    .Value = Before(2).ID 'Number 2 is selected because the array was sorted and the empty element (=0) went
                    'to the first position of the array
                    .Validation.Delete
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:=Join(IntegerToStringArray(Blanks), ",")
                    .Validation.InputMessage = Msg1
                    .Validation.InputTitle = msg2
                    .Validation.ErrorMessage = msg3
                    .Validation.ErrorTitle = msg4
                End With
                
                With SamList_Sh.Range(FindIDObj.Address).Offset(, 4)
                    .Value = After(2).ID 'Number 2 is selected because the array was sorted and the empty element (=0) went
                    'to the first position of the array
                    .Validation.Delete
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(IntegerToStringArray(Blanks), ",")
                    .Validation.InputMessage = Msg1
                    .Validation.InputTitle = msg2
                    .Validation.ErrorMessage = msg3
                    .Validation.ErrorTitle = msg4
                End With
        End With
        
    'This is exactly the same code as above but is used to deal with external standards analyses
        ca = 1
        cb = 1
        
    ReDim Before(1 To 1) As IDsTimesDifference
    ReDim After(1 To 1) As IDsTimesDifference
    
        For Each d In ExtStandards
            B = PathsNamesIDsTimesCycles(4, a) - PathsNamesIDsTimesCycles(4, d)
                If B = 0 Then
                    MsgBox PathsNamesIDsTimesCycles(3, d) & "time of analyses is exactly the same of " & PathsNamesIDsTimesCycles(3, a) & _
                    ". Please, check this and the retry"
                ElseIf B < 0 Then
                    After(ca).ID = d 'ID of the external standard analyses
                    After(ca).TimeDifference = Abs(B) 'Difference of time of analyses between external standard and sample
                        ReDim Preserve After(1 To UBound(After) + 1)
                            ca = ca + 1
                Else
                    Before(cb).ID = d 'ID of the external standardsd analyses
                    Before(cb).TimeDifference = B 'Difference of time of analyses between external standard and sample
                        ReDim Preserve Before(1 To UBound(Before) + 1)
                            cb = cb + 1
                End If

        Next
        
        'There must be a primary standard analysis before and after the sample and this will be checked below
        If After(1).ID = 0 Then
            
            MsgBox "After sample " & PathsNamesIDsTimesCycles(2, a) & " there is not any " & _
            "primary standard analysis. Please, check this " & _
            "missing analysis, its name might be different than the one you selected " & _
            "(" & ExternalStandardName_UPb & ").", vbOKOnly
                
                Call FormatMainSh
                    Call UnloadAll
                        SamList_Sh.Activate
                            End
                        
        ElseIf Before(1).ID = 0 Then
            
            MsgBox "Before sample " & PathsNamesIDsTimesCycles(2, a) & " there is not any " & _
            "primary standard analysis. Please, check this " & _
            "missing analysis, its name might be different than the one you selected " & _
            "(" & ExternalStandardName_UPb & ").", vbOKOnly
                
                Call FormatMainSh
                    Call UnloadAll
                        SamList_Sh.Activate
                            End

        End If
        
        'Sub to sort ascending the time differences
        Call SortMyData(After)
        Call SortMyData(Before)

        With SamList_Sh.Columns(StdSlpColumnNum)
            Set FindIDObj = .Find(a)
            
                'There are some lines below that add valitadion to SamListMap cells (columns I to J).
                
                With SamList_Sh.Range(FindIDObj.Address).Offset(, 1)
                    .Value = Before(2).ID
                    .Validation.Delete
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlEqual, Formula1:=Join(IntegerToStringArray(ExtStandards), ",")
                    .Validation.InputMessage = Msg1
                    .Validation.InputTitle = msg2
                    .Validation.ErrorMessage = msg3
                    .Validation.ErrorTitle = msg4
                End With
                
                With SamList_Sh.Range(FindIDObj.Address).Offset(, 2)
                    .Value = After(2).ID
                    .Validation.Delete
                    .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Join(IntegerToStringArray(ExtStandards), ",")
                    .Validation.InputMessage = Msg1
                    .Validation.InputTitle = msg2
                    .Validation.ErrorMessage = msg3
                    .Validation.ErrorTitle = msg4
                End With
        End With
        
    Next
                
    
End Sub
Sub SetPathsNamesIDsTimesCycles()

    'Proocedure that creates the array PathsIDsTimes, which stores the file pathts, IDs and time of
    'first cycle of every sample
    
    Dim a As Variant
    Dim cell As Range
    
    If mwbk Is Nothing Then 'We need some public variables, so we must be shure that they were set
        Call PublicVariables
    End If

        
    'The three lines below are used to create the PathsIDsTimes array
    'RawDataFilesPaths = 1: ID = 2: TimeFirstCycle = 3: Cycles = 4
    RawDataFilesPaths = 1: FileName = 2: ID = 3: TimeFirstCycle = 4: Cycles = 5

    Set IDsRange = SamList_Sh.Range("C" & SamList_FirstLine, SamList_Sh.Range("C" & SamList_FirstLine).End(xlDown))
    
    a = SamList_HeadersLine2 'Row of headers in SamList_Sh
    
    ReDim Preserve PathsNamesIDsTimesCycles(1 To 5, a - 1 To a - 1) As Variant
    
    'Below we populate PathsIDsTimes array with information of all analyses (samples, internal standards, external standards
    'and blanks. This is necessary so that we can reduce data after running this sub.
    For Each cell In IDsRange
    
    'Columns A and B must have strings. Columns C to J must be filled with only numbers (IDs) and any of its cells
    'must not be empty. So, we check these conditions below.
        
        If IsEmpty(SamList_Sh.Range(SamList_FilePath & a + 1)) = True Then
            Application.Goto SamList_Sh.Range(SamList_FilePath & a + 1)
            GoTo ErrHandler
            ElseIf IsEmpty(SamList_Sh.Range(SamList_ID & a + 1)) = True Or WorksheetFunction.IsNumber(SamList_Sh.Range(SamList_ID & a + 1)) = False Then
                Application.Goto SamList_Sh.Range(SamList_ID & a + 1)
                GoTo ErrHandler
                ElseIf IsEmpty(SamList_Sh.Range(SamList_FirstCycleTime & a + 1)) = True Or WorksheetFunction.IsNumber(SamList_Sh.Range(SamList_FirstCycleTime & a + 1)) = False Then
                    Application.Goto SamList_Sh.Range(SamList_FirstCycleTime & a + 1)
                    GoTo ErrHandler
                    ElseIf IsEmpty(SamList_Sh.Range("E" & a + 1)) = True Then
                        Application.Goto SamList_Sh.Range(SamList_Cycles & a + 1)
                        GoTo ErrHandler

        End If
               
        'Below, we make a copy of the IDs to PathsIDsTimes so that we can access them easily during data reduction.
        PathsNamesIDsTimesCycles(RawDataFilesPaths, a - 1) = SamList_Sh.Range(SamList_FilePath & a + 1).Value
        PathsNamesIDsTimesCycles(FileName, a - 1) = SamList_Sh.Range(SamList_FileName & a + 1).Value
        
        
        PathsNamesIDsTimesCycles(ID, a - 1) = SamList_Sh.Range(SamList_ID & a + 1).Value
        PathsNamesIDsTimesCycles(TimeFirstCycle, a - 1) = SamList_Sh.Range(SamList_FirstCycleTime & a + 1).Value
        PathsNamesIDsTimesCycles(Cycles, a - 1) = SamList_Sh.Range(SamList_Cycles & a + 1).Value
        
        ReDim Preserve PathsNamesIDsTimesCycles(1 To 5, 1 To UBound(PathsNamesIDsTimesCycles, 2) + 1) As Variant
        
        a = a + 1
        
    Next
    
    Exit Sub
    
ErrHandler:
    MsgBox "In SamList sheet there are one or more cells of columns A to C, starting from line 3, that are empty. Please, check them."
        End


End Sub

Sub SortMyData(ByRef Data() As IDsTimesDifference)

    'This beautiful code below aws a great suggestion from Steve Jorgensen in
    'https://stackoverflow.com/questions/4873182/sorting-a-multidimensionnal-array-in-vba
    'http://stackoverflow.com/a/4955848/2449724
    'Below there is a copy of Steve's comment about it.
    'The hard part is that VBA provides no straightforward way to swap rows in a 2D array.
    'For each swap, you're going to have to loop over 5 elements and swap each one, which
    'will be very inefficient. I'm guessing that a 2D array is really not what you should
    'be using anyway though. Does each column have a specific meaning? If so, should you
    'not be using an array of a user-defined type, or an array of objects that are instances
    'of a class module? Even if the 5 columns don't have specific meanings, you could still
    'do this, but define the UDT or class module to have just a single member that is a 5-element
    'array.
    'For the sort algorithm itself, I would use a plain ol' Insertion Sort. 1000 items is
    'actually not that big, and you probably won't notice the difference between an Insertion
    'Sort and Quick Sort, so long as we've made sure that each swap will not be too slow. If
    'you do use a Quick Sort, you'll need to code it carefully to make sure you won't run out
    'of stack space, which can be done, but it's complicated, and Quick Sort is tricky enough
    'already.
    
    'So assuming you use an array of UDTs, and assuming the UDT contains variants named Field1
    'through Field5, and assuming we want to sort on Field2 (for example), then the code might
    'look something like this...
    
    
    
    'The code below is mostly Steve's code with minor differences. I defined a UDT called
    'IDsTimesDifference which is used in this code.
    
    Dim FirstIdx As Long, LastIdx As Long
    FirstIdx = LBound(Data)
    LastIdx = UBound(Data)

    Dim i As Long, J As Long, TEMP As IDsTimesDifference
    For i = FirstIdx To LastIdx - 1
        For J = i + 1 To LastIdx
            If Data(i).TimeDifference > Data(J).TimeDifference Then
                TEMP = Data(i)
                Data(i) = Data(J)
                Data(J) = TEMP
            End If
        Next J
    Next i
End Sub

Sub IdentifyFileType()

'Sub capable of identify if the file is a blank, a sample or a standard, based on the names given by
'the user and then save the cell addresses in SamList_Sh of them.

'Updated 27082015 - The user can choose more than one name for blanks, samples and standards.
'UPDATED 13102015 - A new function that allows the user to ignore some the analysis in the list.
    
    Dim FileNamesStart As Range
    Dim FileNamesEnd As Range
    Dim a As Integer
    Dim C As Variant
    Dim Message1 As Integer
    Dim Message2 As Object
    Dim d(1 To 4) As String
    Dim B As String
    Dim E As Integer
    
    Dim Blk_Names() As String
    Dim ExtStd_Names() As String
    Dim Slp_Names() As String
    Dim IntStd_Names() As String
    Dim All_Names() As String
    
    If FolderPath_UPb Is Nothing Then 'We need some public variables, so we must be shure that they were set
        Call PublicVariables
    End If
    
    Set FileNamesStart = SamList_Sh.Range(SamList_FileName & SamList_FirstLine)
    Set FileNamesEnd = SamList_Sh.Range(FileNamesStart, FileNamesStart.End(xlDown)) 'First cell of the range with the file names
    
    Blk_Names = Split(BlankName_UPb, ",") 'Splits the string with n names in a array with n elements
    ExtStd_Names = Split(ExternalStandardName_UPb, ",") 'Splits the string with n names in a array with n elements
    Slp_Names = Split(SamplesNames_UPb, ",") 'Splits the string with n names in a array with n elements
    IntStd_Names = Split(InternalStandard_UPb, ",") 'Splits the string with n names in a array with n elements
    
    C = ConcatenateArrays(All_Names, Blk_Names)
        C = ConcatenateArrays(All_Names, ExtStd_Names)
            C = ConcatenateArrays(All_Names, Slp_Names)
                If InternalStandardCheck_UPb = True Then
                    C = ConcatenateArrays(All_Names, IntStd_Names)
                End If
    
    For a = LBound(All_Names) To UBound(All_Names)
        All_Names(a) = Replace(All_Names(a), " ", "") 'Removes any spaces from the array elements (all analyses names)
    Next
    
    For a = LBound(All_Names) To UBound(All_Names)
        For E = a + 1 To UBound(All_Names)
            If All_Names(a) = All_Names(E) Then
                MsgBox "Names of samples, blanks and standards are duplicated. Please, check them and then retry."
                    Application.Goto SamplesNames_UPb
                        Call UnloadAll
                            End
            End If
        Next
    Next
    
'    BlkFound = FindStrings(BlankName_UPb.Value, FileNamesStart, FileNamesEnd) 'Cell addresses in column B of blanks
    'SlpFound = FindStrings(SamplesNames_UPb.Value, FileNamesStart, FileNamesEnd) 'Cell addresses in column B of samples
'    StdFound = FindStrings(ExternalStandardName_UPb.Value, FileNamesStart, FileNamesEnd) 'Cell addresses in column B of external standards

    'The code below will create the BlkFound array. This is necessary because more than one name is acceptable for blanks
        For a = LBound(Blk_Names) To UBound(Blk_Names)
            Blk_Names(a) = Replace(Blk_Names(a), " ", "") 'Removes any spaces from the array elements (sample names)
        Next
        
        For a = LBound(Blk_Names) To UBound(Blk_Names)
            'By using "FindStrings", cell addresses in column B of samples are found.
            'By using concatenateArrays, arrays with cell address of each samples are
            'concatenated into only one array.
            C = ConcatenateArrays( _
            BlkFound, FindStrings(Blk_Names(a), _
            FileNamesStart, FileNamesEnd))
        Next


        'The code below will create the SlpFound array. This is necessary because more than one name is acceptable for samples
            For a = LBound(Slp_Names) To UBound(Slp_Names)
                Slp_Names(a) = Replace(Slp_Names(a), " ", "")
            Next
            
            For a = LBound(Slp_Names) To UBound(Slp_Names)
        
                C = ConcatenateArrays( _
                SlpFound, FindStrings(Slp_Names(a), _
                FileNamesStart, FileNamesEnd))
            Next
        
           'The code below will create the IntStdFound array. This is necessary because more than one name is acceptable for internal standards.
               If InternalStandardCheck_UPb = True Then
                   
                       For a = LBound(IntStd_Names) To UBound(IntStd_Names)
                           IntStd_Names(a) = Replace(IntStd_Names(a), " ", "")
                       Next
                       
                       For a = LBound(IntStd_Names) To UBound(IntStd_Names)
                           C = ConcatenateArrays( _
                           IntStdFound, FindStrings(IntStd_Names(a), _
                           FileNamesStart, FileNamesEnd))
                       Next
                   
                   Else
                   
                       IntStdFound = FindStrings(InternalStandard_UPb.Value, FileNamesStart, FileNamesEnd)
               End If
               
               'The code below will create the StdFound array. This is necessary because more than one name is acceptable for primary standards
                   For a = LBound(ExtStd_Names) To UBound(ExtStd_Names)
                       ExtStd_Names(a) = Replace(ExtStd_Names(a), " ", "")
                   Next
                   
                   For a = LBound(ExtStd_Names) To UBound(ExtStd_Names)
                       C = ConcatenateArrays( _
                       StdFound, FindStrings(ExtStd_Names(a), _
                       FileNamesStart, FileNamesEnd))
                   Next
    
    'The lines below will take the arrays of samples, standards and blanks found and test if some of them
    'must be ignored.
    IgnoredAnalysis (BlkFound)
    IgnoredAnalysis (SlpFound)
    IgnoredAnalysis (IntStdFound)
    IgnoredAnalysis (StdFound)

    If IsArrayEmpty(BlkFound) = True Then
        MsgBox "No blanks were found in " & FolderPath_UPb & ". Please, check their names and their files paths.", vbOKOnly
            Application.Goto BlankName_UPb
                Call UnloadAll: End
                
        ElseIf IsArrayEmpty(SlpFound) = True Then
            MsgBox "No samples were found in " & FolderPath_UPb & ". Please, check their names and their files paths.", vbOKOnly
                Application.Goto SamplesNames_UPb
                    Call UnloadAll: End
    
            ElseIf IsArrayEmpty(StdFound) = True Then
                MsgBox "No external standards were found in " & FolderPath_UPb & ". Please, check their names and their files paths.", vbOKOnly
                    Application.Goto ExternalStandardName_UPb
                        Call UnloadAll: End
              
                ElseIf InternalStandardCheck_UPb = True And IsArrayEmpty(IntStdFound) = True Then
                    MsgBox "No Internal standards were found in " & FolderPath_UPb & ". Please, check their names and their files paths.", vbOKOnly
                        Application.Goto InternalStandard_UPb
                            Call UnloadAll: End
    End If
  
End Sub

Sub IgnoredAnalysis(Arr1 As Variant)
    
    'This procedures takes the array of samples standards and blanks analysis found in SmaList sheet and it does a test.
    'The symbol * next to the name of the analysis means that it must be ignored, so the analysis is removed from the
    'array being tested.
    
    'Created 13102015
    
    Dim Counter1 As Long
    Dim DeleteElement As Boolean
    
    If IsArrayEmpty(Arr1) = False Then
        For Counter1 = LBound(Arr1) To UBound(Arr1)
            If Left(SamList_Sh.Range(Arr1(Counter1)), 1) = IgnoreSymbol Then
                DeleteElement = DeleteArrayElement(Arr1, Counter1, True)
            End If
        Next
    End If
    
End Sub

Public Sub IgnoreAnalysis()

    'Created 18122015
    'This procedure ONLY WORKS IN PLOT SHEETS
    
    
End Sub

Public Sub LoadSamListMap()

    'This program recognizes the samples, standards and blanks IDs in SamList, creating an array (AnalysesList which is
    'equal to the SamListMap)that will be used to set the sequence in which data will be reduced.
    
    'THIS SUB IS NOT BEING USED BECAUSE IT NECESSARY TO UPDATE OTHER CHRONUS PROCEDURES. BY DEFAULT,
    'THE LIST OF SAMPLES IS CREATED AGAIN BY MARCOFOLDEROFFICE2010.
            
    Dim cell As Range
    Dim a As Integer
    Dim Counter As Integer
    
    If SampleName_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) = True Then
        Call SetPathsNamesIDsTimesCycles
    End If
    
    Set MapIDsRange = SamList_Sh.Range("H" & SamList_FirstLine, SamList_Sh.Range("H" & SamList_FirstLine).End(xlDown))
        
    a = SamList_FirstLine 'Row number with headers of SamList map
    Counter = 1
    
    ReDim Preserve AnalysesList(1 To 1) As SamplesMap '''''''''''''''''''''''''CHECAR   <<<<<-------------------------------------
    
    For Each cell In MapIDsRange
            
    'The "map" of samples, standards and blanks IDs must be filled with only numbers and any of its cells
    'must no be empty. So, we check these conditions below.
        
        If IsEmpty(SamList_Sh.Range("H" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("H" & a)) = False Then
            Application.Goto SamList_Sh.Range("H" & a)
            GoTo ErrHandler
            ElseIf IsEmpty(SamList_Sh.Range("I" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("I" & a)) = False Then
                Application.Goto SamList_Sh.Range("I" & a)
                GoTo ErrHandler
                ElseIf IsEmpty(SamList_Sh.Range("J" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("J" & a)) = False Then
                    Application.Goto SamList_Sh.Range("J" & a)
                    GoTo ErrHandler
                    ElseIf IsEmpty(SamList_Sh.Range("K" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("K" & a)) = False Then
                        Application.Goto SamList_Sh.Range("K" & a)
                        GoTo ErrHandler
                        ElseIf IsEmpty(SamList_Sh.Range("L" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("L" & a)) = False Then
                            Application.Goto SamList_Sh.Range("L" & a)
                            GoTo ErrHandler
        
        End If
        
        On Error GoTo ErrHandler
        
        'Below, we make a copy of the IDs to AnalysesList so that we can access them easily during data reduction.
        AnalysesList(Counter).sample = SamList_Sh.Range("H" & a).Value
        AnalysesList(Counter).Std1 = SamList_Sh.Range("I" & a).Value
        AnalysesList(Counter).Std2 = SamList_Sh.Range("J" & a).Value
        AnalysesList(Counter).Blk1 = SamList_Sh.Range("K" & a).Value
        AnalysesList(Counter).Blk2 = SamList_Sh.Range("L" & a).Value
        
        ReDim Preserve AnalysesList(1 To UBound(AnalysesList) + 1) As SamplesMap
        
        a = a + 1
        Counter = Counter + 1
        
    Next
    
    Exit Sub

ErrHandler:
    MsgBox "In SamList sheet there are one or more cells of columns H to L, starting from line " & _
        SamList_FirstLine + 1 & ", that are not filled with numbers or empty. Please, check them."
        End
    
    
End Sub

Public Sub LoadStdListMap()

    'This program recognizes the external standards and blanks IDs in SamList, creating an array (AnalysesList_std which is
    'equal to the StdListMap) that will be used to set the sequence in which data will be reduced.
            
    Dim cell As Range
    Dim a As Integer
    Dim Counter As Integer
    
    If SampleName_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) = True Then
        Call SetPathsNamesIDsTimesCycles
    End If
    
    Set MapIDsRange = SamList_Sh.Range("F" & SamList_FirstLine, SamList_Sh.Range("F" & SamList_FirstLine).End(xlDown))
        
    a = SamList_FirstLine 'Row number with headers of SamList map
    Counter = 1
    
    ReDim Preserve AnalysesList_std(1 To 1) As ExtStandardsMap '''''''''''''''''''''''''CHECAR   <<<<<-------------------------------------
    
    For Each cell In MapIDsRange
            
    'The "map" of samples, standards and blanks IDs must be filled with only numbers and any of its cells
    'must no be empty. So, we check these conditions below.
        
        If IsEmpty(SamList_Sh.Range("F" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("F" & a)) = False Then
            Application.Goto SamList_Sh.Range("F" & a)
            GoTo ErrHandler
            ElseIf IsEmpty(SamList_Sh.Range("G" & a)) Or WorksheetFunction.IsNumber(SamList_Sh.Range("G" & a)) = False Then
                Application.Goto SamList_Sh.Range("I" & a)
                GoTo ErrHandler
        End If
        
        On Error GoTo ErrHandler
        
        'Below, we make a copy of the IDs to AnalysesList so that we can access them easily during data reduction.
        AnalysesList_std(Counter).Std = SamList_Sh.Range("F" & a).Value
        AnalysesList_std(Counter).Blk1 = SamList_Sh.Range("G" & a).Value
        
        ReDim Preserve AnalysesList_std(1 To UBound(AnalysesList_std) + 1) As ExtStandardsMap
        
        a = a + 1
        Counter = Counter + 1
        
    Next
    
    Exit Sub

ErrHandler:
    MsgBox "In SamList sheet there are one or more cells of columns H to L, starting from line " & SamList_FirstLine + 1 & ", that are not filled with numbers or empty. Please, check them. " & _
    "Column E is an exception and must be empty."
        End
    
    
End Sub


Sub ClearCycles(WB As Workbook, ChoosenCycles As Variant)
    'This program is supposed to, based on an array with the selected cycles
    'stored in SamList_Sh (column E), clear all other cycles that should not be used.
    'WB is the workbook of the raw data file and ChoosenCycles is the range where the
    'selected cycles for this data are stored (SamList_Sh, column E).
    
    ''Application.DisplayAlerts = False
    
    Dim ChoosenCyclesArray() As String
    Dim AllCycles() As Integer
    Dim NumberCycles As Integer
    Dim a As Variant
    Dim B As Variant
    Dim C As Boolean
    Dim d As Integer
    Dim Counter As Integer
    
    ChoosenCyclesArray = Split(ChoosenCycles, ",")
    NumberCycles = RawNumberCycles_UPb
        If UBound(ChoosenCyclesArray) = 0 Then
            MsgBox "It's impossible to evaluate an analysis with only one cycle. " _
                & "Please, check the cycles that must be considered for " & WB.Name & ". Look at column E."
                    WB.Close savechanges:=False
                        Application.Goto SamList_Sh.Range("A1")
                            End
        End If
        
    ReDim AllCycles(1 To NumberCycles) As Integer
    
    Counter = 1
    For Each a In AllCycles
        AllCycles(Counter) = Counter
            Counter = Counter + 1
    Next
        
    For Each a In AllCycles 'ChoosenCyclesArray
        C = False
        
            For Each B In ChoosenCyclesArray
                If Val(B) > UBound(AllCycles) Then
                    MsgBox "You have choosen an cycle for " & WB.Name & " that doesn't exist. Please, check it."
                        WB.Close savechanges:=False
                            Application.Goto SamList_Sh.Range("A1")
                                End
                End If
                    
                
                If a = Val(B) Then
                    C = True: GoTo 1
                End If
            Next
        
        If C = False Then
            WB.Worksheets(1).Range(RawCyclesTimeRange).Item(a).EntireRow.ClearContents
        End If
1    Next
       
    ''Application.DisplayAlerts = true
        
End Sub

Sub CyclesTime(CyclesTimeRange As Range)
    'This function takes the range of time of analyses of each cycle and,
    'based on cycle duration, changes the values of CyclesTimeRange by
    'multiplying CyclesDuration 1 by the index of the cycle. So we expect
    'that the CyclesTimeRange becomes something like 0, 1, 2, 3,..., until
    'close to the number of cycles.
    
    Dim a As Integer
    
    For a = 1 To RawNumberCycles_UPb
        If Not CyclesTimeRange.Item(a) = "" Then
            CyclesTimeRange.Item(a) = CycleDuration_UPb * a
        End If
    Next
            
End Sub

Sub OpenFilesByIDs()

    Dim a As Range
    
    Call SetPathsNamesIDsTimesCycles
    
    For Each a In Selection
    
        If WorksheetFunction.IsNumber(a.Value) = True Then
        
        On Error Resume Next
            Workbooks.Open (PathsNamesIDsTimesCycles(1, a))
                If Err.Number <> 0 Then
                    MsgBox MissingFile1 & PathsNamesIDsTimesCycles(1, a) & MissingFile2
                        Call UpdateFilesAddresses
                            Call UnloadAll
                                End
                End If
        On Error GoTo 0

            
        End If
    Next
        
End Sub

Sub CreateResultsSh()

    If mwbk Is Nothing Then
        Call PublicVariables
    End If
    
    Set Results_Sh = mwbk.Worksheets.Add
        
        On Error Resume Next
            Results_Sh.Name = "Results_UPb"
                
            If Err.Number <> 0 Then
                If MsgBox("There is a result sheet. Would you like to overwrite it?", vbYesNo) = vbYes Then
                    
                    Results_Sh.Delete
                        Set Results_Sh = mwbk.Worksheets("Results_UPb")
                            Results_Sh.Cells.ClearContents
                            
                Else
                
                    MsgBox "Change the name of the results sheet, if you want to preserve " _
                        & "it and then try again", vbOKOnly
                            On Error GoTo 0
                                Exit Sub
                End If
            End If
        On Error GoTo 0
    
    
End Sub

Sub CopyToCovarSheet(RngArray As Variant)
    
'    Dim a As Variant
'    Dim counter As Integer
'
'    counter = 1
'
'    For Each a In RngArray
'
'        a.Copy Destination:=CovarSheet.Range(Cells(1, counter))
'            counter = counter + 1
'
'    Next
    
End Sub

Sub MatchValidRangeItems(ByVal Rng1 As Range, ByVal Rng2 As Range, ByVal Rng3 As Range, Sh As Worksheet, CalcFirstCell As Range) 'UPDATE DESCRIPTION

    'This procedure will take all the items from both ranges and check if the pairs
    'are valid. By valid I mean they are numeric and not equal to 0. This is
    'fundamental for the ratio calculations where makes no sense at all isotope
    'signal equal to 0 or empty. Pairs that don't pass this test will be deleted.
    'A new range with only valid entries will be set and pasted to the range where
    'CAlcFirstCell is the first cell.
    
    'Rng1 and Rng2 MUST have the same number of rows and only one area.
    'Sh is the parent object of the ranges Rng1 and Rng2
    
    'Returns a range with only valid pairs of cells.
    
    Dim Counter As Long
    Dim Counter2 As Long
    Dim NotValidPairItem() As Integer
    Dim ItemFrom1 As Variant
    Dim ItemFrom2 As Variant
    Dim ItemFrom3 As Variant
    Dim NumRows As Integer
    Dim IsThereEmptyElementArray As Boolean
    Dim Range1 As Range
    Dim Range2 As Range
    Dim Range3 As Range
    
    ReDim NotValidPairItem(1 To 1) As Integer
    
    If Rng1.Areas.count <> 1 Or Rng2.Areas.count <> 1 Or Rng3.Areas.count <> 1 Then
        MsgBox "Rng1, Rng2 and Rng3 have more than 1 area. Function MatchValidRangeItems failed.", vbOKOnly
            End
    End If
    
    If Rng1.Rows.count <> Rng2.Rows.count Or Rng1.Rows.count <> Rng3.Rows.count Then
        MsgBox "Rng1, Rng2 and Rng3 doesn't have the same number of rows. Function MatchValidRangeItems failed.", vbOKOnly
            End
    End If
    
    NumRows = Rng1.Rows.count
    
    'The lines below will clear the ranges where the three ranges will be pasted
    CalcFirstCell.Columns.ClearContents: CalcFirstCell.Offset(, 1).Columns.ClearContents: CalcFirstCell.Offset(, 2).Columns.ClearContents
        Rng1.Copy Destination:=CalcFirstCell
            Rng2.Copy Destination:=CalcFirstCell.Offset(, 1)
                Rng3.Copy Destination:=CalcFirstCell.Offset(, 2)
        
    'To avoid changing the input ranges, I set new ones to the ranges where rng1, rng2 and rng3 contents were pasted.
    Set Range1 = Sh.Range(CalcFirstCell, CalcFirstCell.Offset(NumRows - 1))
    Set Range2 = Sh.Range(CalcFirstCell.Offset(, 1), CalcFirstCell.Offset(NumRows - 1, 1))
    Set Range3 = Sh.Range(CalcFirstCell.Offset(, 2), CalcFirstCell.Offset(NumRows - 1, 2))
    
'    Rng1.Select
'    Rng2.Select

    'All items from Range1, Range1 and Range1 will be checked.
    For Counter = 1 To NumRows
        
        ItemFrom1 = Range1.Item(Counter)
        ItemFrom2 = Range2.Item(Counter)
        ItemFrom3 = Range3.Item(Counter)
        
        'ItemFrom1, ItemFrom2 and ItemFrom3 must be numeric and different from 0
        If WorksheetFunction.IsNumber(ItemFrom1) = False Or _
           WorksheetFunction.IsNumber(ItemFrom2) = False Or _
           WorksheetFunction.IsNumber(ItemFrom3) = False Or _
           ItemFrom1 = 0 Or _
           ItemFrom2 = 0 Or _
           ItemFrom3 = 0 Then
                
                NotValidPairItem(UBound(NotValidPairItem)) = Counter
                
                    If Not Counter = NumRows Then
                        ReDim Preserve NotValidPairItem(1 To UBound(NotValidPairItem) + 1)
                    End If
        End If
        
    Next
    
'    If IsEmpty(NotValidPairItem(LBound(NotValidPairItem))) Then
'        MsgBox "All pairs of values are not valid.", vbOKOnly
'            Rng1.Select
'                End
'    End If
    If Not NotValidPairItem(LBound(NotValidPairItem)) = 0 Then
        
        If NotValidPairItem(UBound(NotValidPairItem)) = 0 Then 'The last array element will always be empty if the last item from Range1, Range2 and Range3 don't fail the previous test.
            IsThereEmptyElementArray = DeleteArrayElement(NotValidPairItem, UBound(NotValidPairItem), True)
        End If
    
        For Counter = 1 To UBound(NotValidPairItem)
            For Counter2 = 1 To NumRows
                If Counter2 = NotValidPairItem(Counter) Then
                    
                    Range1.Item(Counter2).ClearContents
                    Range2.Item(Counter2).ClearContents
                    Range3.Item(Counter2).ClearContents
                    
                    Counter2 = NumRows
                End If
            Next
        Next
    End If

End Sub

Sub UnloadAll()

    Application.DisplayAlerts = False
    
    'By Reafidy, from http://www.ozgrid.com/forum/showthread.php?t=86401 19012015
    Dim objLoop As Object
     
    For Each objLoop In VBA.UserForms
    'MsgBox objLoop.Name
        If TypeOf objLoop Is UserForm Then Unload objLoop
        
    Next objLoop
    
End Sub

Sub CreateFinalReport()

    Dim PasteRow As Long
    Dim Counter As Long
    Dim Range_SlpStdBlkCorr As Range
    Dim Range_SlpStdCorr As Range
    Dim Range_SlpStdCorrHeaders As Range
    Dim CellRange As Range
    Dim FindID As Object
    Dim ScreenUpdt As Boolean
    Dim LastItem As Long
    Dim SearchStr As Variant
    
    ScreenUpdt = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
        
    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If
                                        
    On Error Resume Next
        Application.DisplayAlerts = False
            Set FinalReport_Sh = mwbk.Sheets(FinalReport_Sh_Name)
                If Err.Number = 0 Then
                    
                    If MsgBox("There is already a report in this workbook. Would you like to overwrite it?", vbYesNo) = vbYes Then
                        
                        mwbk.Sheets(FinalReport_Sh_Name).Delete
                            
                            TW.Sheets(FinalReport_Sh_Name).Copy _
                                After:=mwbk.Sheets(SlpStdCorr_Sh_Name)
                        
                            Set FinalReport_Sh = mwbk.Sheets(FinalReport_Sh_Name)
                            
                    Else
                    
                        Call UnloadAll
                            End
                    
                    End If
                    
                Else
                
                    TW.Sheets(FinalReport_Sh_Name).Copy _
                    After:=mwbk.Sheets(SlpStdCorr_Sh_Name)
                    
                        Set FinalReport_Sh = mwbk.Sheets(FinalReport_Sh_Name)
                
                End If
        Application.DisplayAlerts = True
    On Error GoTo 0
        
    Call SetPathsNamesIDsTimesCycles
    
    With SlpStdCorr_Sh
        .AutoFilterMode = False
        Set Range_SlpStdCorr = .Range(.Range(StdCorr_ColumnID & StdCorr_HeaderRow + 1), .Range(StdCorr_ColumnID & StdCorr_HeaderRow).End(xlDown))
        Set Range_SlpStdCorrHeaders = .Range(.Range(StdCorr_FirstColumn & StdCorr_HeaderRow), .Range(StdCorr_LastColumn & StdCorr_HeaderRow))
    End With
        
        With SlpStdBlkCorr_Sh
            Set Range_SlpStdBlkCorr = .Range(.Range(ColumnID & HeaderRow + 1), .Range(ColumnID & HeaderRow + 1).End(xlDown))
        End With
    
    If Range_SlpStdCorr.Item(1) = "" Or Range_SlpStdBlkCorr.Item(1) = "" Then
        Exit Sub
    End If

    PasteRow = 2
    Counter = PasteRow
    
    LastItem = Range_SlpStdCorr.count

    SearchStr = InStr(SlpStdCorr_Sh.Range(StdCorr_Column681Std & HeaderRow), "(%)")
    
        If WorksheetFunction.IsNumber(SearchStr) = True Then
            If SearchStr <> 0 Then
                Call ConvertAbsolute
            End If
        End If

    
    For Each CellRange In Range_SlpStdCorr

        With SlpStdCorr_Sh
        
            'The three lines below copy the color of the cells in SlpStdCorr_Sh and the strikethrough state
            With FinalReport_Sh.Range(FR_ColumnSlpName & FR_HeaderRow + PasteRow, FR_LastColumn & FR_HeaderRow + PasteRow)
                With .Font
                    .Color = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Font.Color
                    .TintAndShade = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Font.TintAndShade
                    .Strikethrough = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Font.Strikethrough
                End With
                
                With .Interior
                        .Pattern = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Interior.Pattern
                        .PatternColorIndex = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Interior.PatternColorIndex
                        .PatternTintAndShade = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Interior.PatternTintAndShade
                        .ThemeColor = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Interior.ThemeColor
                        .TintAndShade = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellRange.Row).Interior.TintAndShade
                End With
            End With
        
            FinalReport_Sh.Range(FR_ColumnSlpName & FR_HeaderRow + PasteRow) = .Range(StdCorr_SlpName & CellRange.Row)
            FinalReport_Sh.Range(FR_Column204PbCps & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column4 & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnThU & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column28 & CellRange.Row)
            FinalReport_Sh.Range(FR_Column64 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column64 & CellRange.Row)
            FinalReport_Sh.Range(FR_Column641Std & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column641Std & CellRange.Row)
                
                
            FinalReport_Sh.Range(FR_ColumnWeth75 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column75 & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnWeth751Std & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column751Std & CellRange.Row)

            FinalReport_Sh.Range(FR_ColumnWeth68 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column68 & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnWeth681Std & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column681Std & CellRange.Row)
                    
            FinalReport_Sh.Range(FR_ColumnWethRho & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column7568Rho & CellRange.Row)
            
            FinalReport_Sh.Range(FR_ColumnAge68 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column68AgeMa & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnAge682StdAbs & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column68AgeMa1std & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnAge75 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column75AgeMa & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnAge752StdAbs & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column75AgeMa1std & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnAge76 & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column76AgeMa & CellRange.Row)
            FinalReport_Sh.Range(FR_ColumnAge762StdAbs & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column76AgeMa1std & CellRange.Row)
            FinalReport_Sh.Range(FR_Column6876DiscPercent & FR_HeaderRow + PasteRow) = .Range(StdCorr_Column6876Conc & CellRange.Row)
                        
        End With
        
        With SlpStdBlkCorr_Sh
            
            With Range_SlpStdBlkCorr
                
                Set FindID = .Find(CellRange)
                    
                    If Not FindID Is Nothing Then
                    
                        FinalReport_Sh.Range(FR_Column206PbmV & FR_HeaderRow + PasteRow) = SlpStdBlkCorr_Sh.Range(Column6 & FindID.Row)
                            
                            'Converting 206 cps to mV
                            If WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_Column206PbmV & FR_HeaderRow + PasteRow)) = True Then
                                FinalReport_Sh.Range(FR_Column206PbmV & FR_HeaderRow + PasteRow) = _
                                    FinalReport_Sh.Range(FR_Column206PbmV & FR_HeaderRow + PasteRow) / mVtoCPS_UPb
                            End If
                        
                    End If
                    
            End With
            
        End With
        
               'Error 64 will be converted to percentage
               If _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_Column641Std & FR_HeaderRow + PasteRow)) And _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_Column64 & FR_HeaderRow + PasteRow)) Then
                    
                    FinalReport_Sh.Range(FR_Column641Std & FR_HeaderRow + PasteRow) = _
                        100 * ( _
                        FinalReport_Sh.Range(FR_Column641Std & FR_HeaderRow + PasteRow) / _
                        FinalReport_Sh.Range(FR_Column64 & FR_HeaderRow + PasteRow))
                
                End If
                
                'Error 75 will be converted to percentage
                If _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_ColumnWeth751Std & FR_HeaderRow + PasteRow)) And _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_ColumnWeth75 & FR_HeaderRow + PasteRow)) Then
                
                    FinalReport_Sh.Range(FR_ColumnWeth751Std & FR_HeaderRow + PasteRow) = _
                        100 * _
                        FinalReport_Sh.Range(FR_ColumnWeth751Std & FR_HeaderRow + PasteRow) / _
                        FinalReport_Sh.Range(FR_ColumnWeth75 & FR_HeaderRow + PasteRow)

                End If
                
                'Error 68 will be converted to percentage
                If _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_ColumnWeth681Std & FR_HeaderRow + PasteRow)) And _
                    WorksheetFunction.IsNumber(FinalReport_Sh.Range(FR_ColumnWeth68 & FR_HeaderRow + PasteRow)) Then
                
                    FinalReport_Sh.Range(FR_ColumnWeth681Std & FR_HeaderRow + PasteRow) = _
                        100 * _
                        FinalReport_Sh.Range(FR_ColumnWeth681Std & FR_HeaderRow + PasteRow) / _
                        FinalReport_Sh.Range(FR_ColumnWeth68 & FR_HeaderRow + PasteRow)
                End If
                
            'Age uncertainties become 2 stddev by multiplying them by 2
            On Error Resume Next
                FinalReport_Sh.Range(FR_ColumnAge682StdAbs & FR_HeaderRow + PasteRow) = _
                    2 * FinalReport_Sh.Range(FR_ColumnAge682StdAbs & FR_HeaderRow + PasteRow)
                FinalReport_Sh.Range(FR_ColumnAge752StdAbs & FR_HeaderRow + PasteRow) = _
                    2 * FinalReport_Sh.Range(FR_ColumnAge752StdAbs & FR_HeaderRow + PasteRow)
                FinalReport_Sh.Range(FR_ColumnAge762StdAbs & FR_HeaderRow + PasteRow) = _
                    2 * FinalReport_Sh.Range(FR_ColumnAge762StdAbs & FR_HeaderRow + PasteRow)
            On Error GoTo 0
        
        'Copying format
                
        
        If Not CellRange.Row = Range_SlpStdCorr.Item(LastItem).Row Then
            With FinalReport_Sh
                .Rows(FR_HeaderRow + PasteRow).Insert Shift:=xlDown 'Adding a new row
                    .Rows(FR_HeaderRow + Counter + 1).EntireRow.Copy 'Copying a row with thencorrectnumber format
                        .Rows(FR_HeaderRow + PasteRow).PasteSpecial Paste:=xlPasteAllExceptBorders 'pasting to the new row before it receives any valy
                            .Rows(FR_HeaderRow + PasteRow).EntireRow.ClearContents 'Cleaning the new row contents
            End With
        End If
    
        Counter = Counter + 1
    Next
    
    Call FormatFinalReport
    
    Range_SlpStdCorrHeaders.AutoFilter
    
    Call UnloadAll

    Application.ScreenUpdating = ScreenUpdt
    
End Sub

Sub UpdateFilesAddresses()

    Dim Msg1 As Integer
    Dim FileAddresses As Range
    
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
        
    With SamList_Sh
        Set FileAddresses = .Range(.Range(SamList_FilePath & SamList_FirstLine), .Range(SamList_FilePath & SamList_FirstLine).End(xlDown))
    End With
    
    Msg1 = MsgBox("Would you like to update the address of the folder where your files are?", vbYesNo)
    
        If Msg1 = vbYes Then 'This IF structure should be updated to deal with the possibility of the FolderPath_UPb be empty.    <---------
            OldFolderPath = FolderPath_UPb
                If Not OldFolderPath = "" Then
                    Call SelectFolder
                        
                    If Right(NewFolderPath, 1) <> "\" Then 'Adds the "\" to the end of the address
                        NewFolderPath = NewFolderPath & "\"
                    End If
                    
                    FolderPath_UPb = NewFolderPath
                                    
                    SamList_Sh.Cells.Replace _
                        What:=OldFolderPath, _
                        Replacement:=NewFolderPath, _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        MatchCase:=False, _
                        SearchFormat:=False, _
                        ReplaceFormat:=False
                    
                    MsgBox "Folder address updated, now you can try to run Chronus again."
                
                Else
                    MsgBox "It is not possible to update the addresses because the original address was not found"
                        Exit Sub

                End If
        End If
    
End Sub

Sub AskToPreserveCycles()
    
    'Program to ask the user if he/she wants to preserve the cycles previously selected or
    'ignore them and reduce data using all of them
    
    Dim PreserveCyclesMsg As Integer

    PreserveCycles = False
    
    If IsEmpty(SamList_Sh.Range(SamList_Cycles & SamList_FirstLine)) = False Then
    
        PreserveCyclesMsg = MsgBox("Would you like to preserve the cycles selected previously?", vbYesNoCancel)
        
        If PreserveCyclesMsg = vbYes Then
        
            PreserveCycles = True
        
        ElseIf PreserveCyclesMsg = vbCancel Then
            
            Call UnloadAll
        
                End
                
        Else
            
            If MsgBox("This mean that any cycle previously removed will be considered now again.", vbOKCancel) = vbCancel Then
                Call UnloadAll
                    End
            End If
        
        End If
    
    End If

End Sub

Sub AskToPreserveListMaps()

    'Program to ask the user if he/she wants to preserve the SamList and StdList maps, considering
    'that these maps can be changed manually.

    Dim PreserveMapsMsg As Integer
    
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If

    If PreserveMaps = True Then
        Exit Sub
    End If
    
    If IsEmpty(SamList_Sh.Range(SamList_SlpID & SamList_FirstLine)) = False Then
    
        PreserveMapsMsg = MsgBox("Would you like to preserve the blanks indicated to each standard and the blanks " & _
        "and standards to each sample?", vbYesNoCancel)
        
        If PreserveMapsMsg = vbYes Then
        
            PreserveMaps = True
        
        ElseIf PreserveMapsMsg = vbCancel Then
            
            Call UnloadAll
        
                End
                
        Else
            
            If MsgBox("This mean that all standards and blanks will automatically selected by Chronus based on " & _
            "the time of analysis.", vbOKCancel) = vbCancel Then
                
                Call UnloadAll
                    
                    End
            End If
        
        End If
    
    End If

End Sub

Sub ChangeAnalysesDate()
    'This is a prototype of a procedure to change the data of analysis in the
    'analyses to be reduced.

    'Created 17102015
    Dim Book As Workbook
    
    Application.ScreenUpdating = False
    
    For Each Book In Workbooks
    
        Range("A4").FormulaR1C1 = "Date: 19/10/2015"
        
        ActiveWorkbook.Save
        
            Application.DisplayAlerts = False
                ActiveWorkbook.Close
            Application.DisplayAlerts = True
            
    Next
    
    Application.ScreenUpdating = True
End Sub

Sub FullDataReductionNew(Optional Program0 As Boolean = True, Optional Program1 As Boolean = True, _
Optional Program2 As Boolean = True, Optional Program3 As Boolean = True, Optional Program4 As Boolean = True, _
Optional Program5 As Boolean = True, Optional Program6 As Boolean = True, Optional Program7 As Boolean = True)

    Dim StartTime As Double
    Dim StartTimeSpecific As Double
    Dim MainProgramsTime As Double
    
    Dim EndTime As Double
    Dim EndTime0 As Double
    Dim EndTime1 As Double
    Dim EndTime2 As Double
    Dim EndTime3 As Double
    Dim EndTime4 As Double
    Dim EndTime5 As Double
    Dim EndTime6 As Double
    Dim EndTime7 As Double
    Dim EndTime8 As Double
    Dim EndTime9 As Double
    Dim EndTime10 As Double
    Dim EndTime11 As Double
    Dim EndTime12 As Double
    Dim EndTime13 As Double
    Dim EndTime14 As Double
    Dim EndTime15 As Double
    Dim EndTime16 As Double
    Dim EndTime17 As Double
    
    Dim DeltaTime0 As Double
    Dim DeltaTime1 As Double
    Dim DeltaTime2 As Double
    Dim DeltaTime3 As Double
    Dim DeltaTime4 As Double
    Dim DeltaTime5 As Double
    Dim DeltaTime6 As Double
    Dim DeltaTime7 As Double
    Dim DeltaTime8 As Double
    Dim DeltaTime9 As Double
    Dim DeltaTime10 As Double
    Dim DeltaTime11 As Double
    Dim DeltaTime12 As Double
    Dim DeltaTime13 As Double
    Dim DeltaTime14 As Double
    Dim DeltaTime15 As Double
    Dim DeltaTime16 As Double
    Dim DeltaTime17 As Double
    
    Dim TotalAnalyses As Long
    
    On Error GoTo 0
    
    StartTime = Timer
        
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    If IsUserFormLoaded(Box7_FullReduction.Name) = False Then
        MsgBox "Box7_FullReduction not loaded."
    End If
    
    Call PublicVariables
    
    Call Load_UPbStandardsTypeList
    
    Call AskToPreserveCycles
    
    If PreserveCycles = True Then
        Call BackupCycles
    End If
    
    Call unprotectsheets

    Call FormatMainSh
    
    If Program0 = True Then
        
        StartTimeSpecific = Timer
        
        Call MacroFolderOffice2010
            EndTime0 = Timer
                DeltaTime0 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime0
                        
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox0, DeltaTime0)
                        
        mwbk.Save

    End If

    If Program1 = True Then
        
        StartTimeSpecific = Timer
        
        Call CheckRawData
            EndTime1 = Timer
                DeltaTime1 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime1
                    
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox1, DeltaTime1)
    End If
    
    If Program2 = True Then
        
        StartTimeSpecific = Timer
        
        Call FirstCycleTime
            EndTime2 = Timer
                DeltaTime2 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime2
                    
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox2, DeltaTime2)
        
        mwbk.Save
    
    End If
    
    Call IdentifyFileType

    mwbk.Save

    If Program3 = True Then
    
        StartTimeSpecific = Timer
        
        Call CreateStdListMap
            EndTime3 = Timer
                DeltaTime3 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime3
    
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox3, DeltaTime3)
                        
        mwbk.Save

    End If
    
    If Program4 = True Then
        
        StartTimeSpecific = Timer
        
        Call CreateSamListMap
            EndTime4 = Timer
                DeltaTime4 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime4
    
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox4, DeltaTime4)
    
        mwbk.Save

    End If
    
    If Program5 = True Then
        
        StartTimeSpecific = Timer
        
        Call CalcBlank
            EndTime5 = Timer
                DeltaTime5 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime5
    
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox5, DeltaTime5)
    
        mwbk.Save
    
    End If
        
    If Program6 = True Then
        
        StartTimeSpecific = Timer
    
        Call CalcAllSlpStd_BlkCorr
            EndTime6 = Timer
                DeltaTime6 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime6
                        
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox6, DeltaTime6)
        
        mwbk.Save

    End If
    
    If Program7 = True Then
        
        StartTimeSpecific = Timer
        
        Call CalcAllSlp_StdCorr
            EndTime7 = Timer
                DeltaTime7 = Timer - StartTimeSpecific
                    MainProgramsTime = MainProgramsTime + DeltaTime7
        
                        Call UpdateFullReductionForm1(Box7_FullReduction.TextBox7, DeltaTime7)
        
        mwbk.Save
        
    End If
    
    Call protectsheets
    
    If PreserveCycles = True Then
        Call RestoreCycles
    End If
    
    mwbk.Save
    
    EndTime = Timer - StartTime
    
    If AllSamplesPath Is Nothing Then
        Set AllSamplesPath = SamList_Sh.Range("A" & SamList_FirstLine, SamList_Sh.Range("A" & SamList_FirstLine).End(xlDown))
    End If
    
    TotalAnalyses = AllSamplesPath.count
    
    Call UpdateFullReductionForm2( _
            EndTime, _
            TotalAnalyses, _
            MainProgramsTime, _
            DeltaTime0, _
            DeltaTime1, _
            DeltaTime2, _
            DeltaTime3, _
            DeltaTime4, _
            DeltaTime5, _
            DeltaTime6, _
            DeltaTime7, _
            Box7_FullReduction.TextBox0, _
            Box7_FullReduction.TextBox1, _
            Box7_FullReduction.TextBox2, _
            Box7_FullReduction.TextBox3, _
            Box7_FullReduction.TextBox4, _
            Box7_FullReduction.TextBox5, _
            Box7_FullReduction.TextBox6, _
            Box7_FullReduction.TextBox7)
    
    Call FormatMainSh
    
    DoEvents
           
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True


End Sub

Sub BackupCycles()
    
    'This procedure will take the list of selected cycles in Samlist sheet and store in an array. At some point, this array
    'will be accessed so the cycles be reinserted in the samlist sheet.
    
    'Created 23102015
    
    Dim Counter As Long
    Dim CyclesRange As Range
    Dim NamesRange As Range
    Dim NumberSamples1 As Long
    Dim NumberSamples2 As Long
    
    If SamList_Sh Is Nothing Then Call PublicVariables
    
    Set CyclesRange = SamList_Sh.Range(SamList_Cycles & SamList_HeadersLine2 + 1, SamList_Sh.Range(SamList_Cycles & SamList_HeadersLine2 + 1).End(xlDown))
    Set NamesRange = SamList_Sh.Range(SamList_FileName & SamList_HeadersLine2 + 1, SamList_Sh.Range(SamList_FileName & SamList_HeadersLine2 + 1).End(xlDown))
    
    NumberSamples1 = CyclesRange.count
    NumberSamples2 = NamesRange.count
        
        If NumberSamples1 <> NumberSamples2 Then
            MsgBox "Please, check the file names and selected cycles ranges in SamList sheet. It is possible that items are missing from one of these ranges.", vbOKOnly
                Call UnloadAll
                    End
        End If
    
    ReDim CyclesBackUpArr(1 To 2, 1 To NumberSamples1) As String
    
    For Counter = 1 To NumberSamples1
        CyclesBackUpArr(1, Counter) = CyclesRange.Item(Counter)
        CyclesBackUpArr(2, Counter) = NamesRange.Item(Counter)
    Next
    
End Sub

Sub RestoreCycles()

    'Created 23102015
    'This procedure will take the CyclesBackUpArr array, compare the names of the files in this array with those in SamList and then
    'copy the selected cycles in the array to the sheet.
    
    Dim Counter1 As Long
    Dim Cell1 As Range
    Dim NamesRange As Range
    Dim NamesMatched As Boolean
    Dim Foldername As String
    
    
    NamesMatched = False
    
    If IsArrayEmpty(CyclesBackUpArr) = False Then
        
            Set NamesRange = SamList_Sh.Range(SamList_FileName & SamList_HeadersLine2 + 1, SamList_Sh.Range(SamList_FileName & SamList_HeadersLine2 + 1).End(xlDown))
                
                If NamesRange.count < UBound(CyclesBackUpArr, 2) Then
                    If MsgBox("There are less analyses in the selected folder. Would you like to stop " & _
                        "the program and check the folder?", vbYesNo) = vbYes Then
                            
                            Call FormatMainSh
                                Shell "C:\WINDOWS\explorer.exe """ & FolderPath_UPb.Value & "", vbNormalFocus
                                    Call UnloadAll
                                        End
                    Else
                        Call FirstCycleTime 'this procedure because it will open wach data file, copie the time when the first cycle was analyzed
                                            'and then fill cycles range with the standard cycles.
                    
                    End If
                End If
            
                For Each Cell1 In NamesRange
                    
                    For Counter1 = 1 To UBound(CyclesBackUpArr, 2)
                        
                        If Cell1 = CyclesBackUpArr(2, Counter1) Then
                            SamList_Sh.Range(SamList_Cycles & Cell1.Row) = CyclesBackUpArr(1, Counter1)
                                NamesMatched = True
                                    Exit For
                        End If
                    
                    Next
                    
                Next
                
    Else
        MsgBox "CyclesBackUpArr was not properly created."
            Call UnloadAll
                End
    End If
    
    If NamesMatched = False Then
        MsgBox "It was not possible to restore at least one of the cycles previously selected. Any changes in the names of the files was made?" & _
            "The program will stop.", vbOKOnly
            Call UnloadAll
                End
    End If
        
End Sub

Sub UpdateFullReductionForm1(ByRef TxtB As MSForms.TextBox, ByVal DeltaTime As Double)
    
'    If IsUserFormLoaded(Box7_FullReduction.Name) = True Then
    Dim ScrUpdt As Boolean
    
    ScrUpdt = Application.ScreenUpdating
    
    Application.ScreenUpdating = True
        With TxtB
            .Text = Round(DeltaTime, 2)
            .BackColor = vbGreen
            .ForeColor = vbBlack
        End With
    DoEvents
    Application.ScreenUpdating = ScrUpdt
    
'    End If

End Sub

Sub UpdateFullReductionForm2( _
        ByVal EndTime As Double, _
        ByVal TotalAnalyses As Long, _
        ByVal MainProgramsTime As Double, _
        ByVal DeltaTime0 As Double, _
        ByVal DeltaTime1 As Double, _
        ByVal DeltaTime2 As Double, _
        ByVal DeltaTime3 As Double, _
        ByVal DeltaTime4 As Double, _
        ByVal DeltaTime5 As Double, _
        ByVal DeltaTime6 As Double, _
        ByVal DeltaTime7 As Double, _
        ByRef Txt0 As MSForms.TextBox, _
        ByRef Txt1 As MSForms.TextBox, _
        ByRef Txt2 As MSForms.TextBox, _
        ByRef Txt3 As MSForms.TextBox, _
        ByRef Txt4 As MSForms.TextBox, _
        ByRef Txt5 As MSForms.TextBox, _
        ByRef Txt6 As MSForms.TextBox, _
        ByRef Txt7 As MSForms.TextBox)
        
    Dim Ctl As Control
    
    If Box7_FullReduction.Program0.Value = True Then
        Txt0.Text = Round((DeltaTime0 / MainProgramsTime), 2)
    End If

    If Box7_FullReduction.Program1.Value = True Then
        Txt1.Text = Round((DeltaTime1 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program2.Value = True Then
        Txt2.Text = Round((DeltaTime2 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program3.Value = True Then
        Txt3.Text = Round((DeltaTime3 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program4.Value = True Then
        Txt4.Text = Round((DeltaTime4 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program5.Value = True Then
        Txt5.Text = Round((DeltaTime5 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program6.Value = True Then
        Txt6.Text = Round((DeltaTime6 / MainProgramsTime), 2)
    End If
    
    If Box7_FullReduction.Program7.Value = True Then
        Txt7.Text = Round((DeltaTime7 / MainProgramsTime), 2)
    End If
    
    With Box7_FullReduction
        
        On Error Resume Next
            .TextBox9.Text = Round(EndTime, 2)
            .TextBox10.Text = TotalAnalyses
            .TextBox11.Text = Round((EndTime / TotalAnalyses), 2)
        On Error GoTo 0
        
    End With
    
    Box7_FullReduction.CommandButton1.Caption = "COMPLETE!"
    
    Box7_FullReduction.CommandButton2.Visible = True

    For Each Ctl In Box7_FullReduction.Controls
        If TypeName(Ctl) = "TextBox" Then
            Ctl.BackColor = vbGreen
            Ctl.ForeColor = vbBlack
        End If
    Next

End Sub

