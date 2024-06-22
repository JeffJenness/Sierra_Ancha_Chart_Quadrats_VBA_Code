# Sierra_Ancha_Chart_Quadrats_VBA_Code
VBA code to create data paper "Cover and density of semi-arid desert grassland patches within Arizona interior chaparral in permanent chart quadrats (1935-2023)"

This repository contains VBA and ArcObjects code used to analyze plant distributions in digitized quadrats near Sierra Ancha Arizona, intermittently over the years 1935 - 1955, and then yearly from 2017 - 2023. This code was used to produce the original data presented in the Data Paper "Cover and density of semi-arid desert grassland patches within Arizona interior chaparral in permanent chart quadrats (1935-2023)" (Moore et al. 2024; see also https://doi.org/...tbd...).

The relevant functions are embedded in larger modules containing other unused functions (14 VBA modules containing 747 functions and 72,608 lines of code). The primary analytical master function is "RunAsBatch" in the module "SierraAnchaAnalysis". This function runs several other functions that do the various steps of the analysis. In particular:

The function "OrganizeData_SA" in module "SierraAnchaAnalysis" assembles all original datasets into a single workspace with a common naming convention, and adds verbatim fields to keep track of edits made to data.
The function "ReviseShapefiles_SA" in module "SierraAnchaAnalysis" corrects species misspellings and misidentifications.
The function "ConvertPointShapefiles_SA" in module "SierraAnchaAnalysis" converts point features to small polygons, deletes a few extraneous objects, adds a few observations that were missed in the digitizing, switches species designations from Cover to Density or vice-versa if necessary, and rotates quadrats if they were mapped with the wrong orientation.  This function calls on data read from Excel files "Rotation.xlsx and "Sierra_Ancha_Species_Lists_Aug 2021_MMM_jsj.xlsx"
The function "AddEmptyFeaturesAndFeatureClasses_SA" in module "SierraAnchaAnalysis" adds empty feature classes if a survey was done on that quadrat in that year but no features were found.  These empty feature classes distinguish these cases from times when no survey was conducted.
The function "RepairOverlappingPolygons_SA" in the module "More_Margaret_Functions" fixes cases when polygons for a single observation are digitized twice, or when separate polygons for a single species overlap.
The function "RecreateSubsetsOfConvertedDatasets_SA" in the module "More_Margaret_Functions" combines all newly-corrected feature classes into a new workspace, and creates two global feature classes containing all cover and all density observations.
The function "AddEmptyFeaturesAndFeatureClassesToCleaned_SA" in module "SierraAnchaAnalysis" adds empty feature classes to the newly corrected feature classes if a survey was done on that quadrat in that year but no features were found.  These empty feature classes distinguish these cases from times when no survey was conducted.
The function "ShiftFinishedShapefilesToCoordinateSystem_SA" in module "SierraAnchaAnalysis" correctly georeferences all feature classes and saves to a new workspace.  Prior to this step all plant locations were in a local 1-square-meter coordinate system based on the 1-square-meter quadrat.
The function "ExportFinalDataset_SA" in module "SierraAnchaAnalysis" removes extraneous and verbatim fields, and exports the final version of the dataset to a new workspace.
The function "SummarizeSpeciesBySite_SA" in module "More_Margaret_Functions" analyzes all the feature classes to determine which species were observed at each site.
The function "SummarizeSpeciesByCorrectQuadrat_SA" in module "More_Margaret_Functions" analyzes all the feature classes to determine which species were observed at each quadrat.
The function "SummarizeYearByCorrectQuadratByYear_SA" in module "More_Margaret_Functions" analyzes all the feature classes to determine which quadrats were surveyed each year.
The function "ExportSubsetsOfSpeciesShapefiles_SA" in module "Margaret_Functions_3" extracts each species individually from the full dataset, and saves them in a series of nested folders suitable for Integral Projection Model functions in R.
The function "CreateFinalTables_SA" in module "SierraAnchaAnalysis" produces the final summary tables intended for distribution with the data, including a list of plant species observed, a summary of the basal area per species by quadrat and year, summary data describing all quadrats and overstory plots, and tabular versions of the global cover and density feature classes.
The function "GenerateRData" in module "SierraAncha_Compare" produces a single CSV file enumerating the presence or absence of all plant species from all quadrats for all years, formatted for a specific analysis in R.

The primary map export function is "ExportImages_SA" in the module "Margaret_Maps", and is run separately from the 15 functions run in the batch file above. This map-making function creates common plant species symbology that can be applied to all 312 maps, and exports individual maps for each quadrat and for each year. This function is best run from an ArcMap document with no data in it, which is why it is run separately from the other functions.

Logistical and financial support was provided by US Forest Service Rocky Mountain Research Station to Northern Arizona University (NAU) through a Research Joint Venture Agreement and by NAU Ecological Restoration institute in-kind staff and undergraduate personnel support.  Financial support was also provided by the National Science Foundation, grant DEB-1906243 and NAU Office of Vice President of Research grants for UG mentored research.

Financial support was also provided by the National Research Initiative of the USDA Cooperative State Research, Education and Extension Service, grant number 2003-35101-12919 and by the National Science Foundation, grant DEB-1906243.