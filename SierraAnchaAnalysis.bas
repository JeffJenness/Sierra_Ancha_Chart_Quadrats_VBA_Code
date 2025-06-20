Attribute VB_Name = "SierraAnchaAnalysis"
Option Explicit

Public Sub RunAsBatch()
Attribute RunAsBatch.VB_ProcData.VB_Invoke_Func = ""

  Dim lngTimeStart As Long
  lngTimeStart = GetTickCount

  More_Margaret_Functions.SummarizeSpeciesBySite_SA
  More_Margaret_Functions.SummarizeSpeciesByCorrectQuadrat_SA
  More_Margaret_Functions.SummarizeYearByCorrectQuadratByYear_SA
  SierraAncha_Compare.GenerateRData

  CreateFinalTables_SA

  Debug.Print "============================"
  Debug.Print "Batch Process Complete:"
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngTimeStart)
End Sub

Public Function ReturnSpeciesTypeColl(pCoverFClass As IFeatureClass, pDensityFClass As IFeatureClass, pSpeciesData As Collection, _
      Optional pRefColl_Apr19 As Collection, Optional strFinalFolder As String) As Collection

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Summary_Data_from_JSJ\Master_List_Species_Apr_19_2023_WG.xlsx", 0)

  Dim pReturn As New Collection
  Dim pTable As ITable
  Dim pRow As IRow
  Dim pCursor As ICursor
  Dim strSpecies As String
  Dim strType As String
  Dim lngCoverIndex As Long
  Dim lngDensityIndex As Long

  Set pRefColl_Apr19 = New Collection
  Set pTable = pWS.OpenTable("Sheet1$")
  lngCoverIndex = pTable.FindField("Cover_Species")
  lngDensityIndex = pTable.FindField("Density_Species")

  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    If Not IsNull(pRow.Value(lngCoverIndex)) Then
      strSpecies = Trim(pRow.Value(lngCoverIndex))
      If strSpecies <> "" Then pRefColl_Apr19.Add "Cover", strSpecies
    End If
    If Not IsNull(pRow.Value(lngDensityIndex)) Then
      strSpecies = Trim(pRow.Value(lngDensityIndex))
      If strSpecies <> "" Then pRefColl_Apr19.Add "Density", strSpecies
    End If
    Set pRow = pCursor.NextRow
  Loop

  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Summary_Data_from_JSJ\Sierra Ancha_CD_Plant Species List_Master_2023.xlsx", 0)
  Set pSpeciesData = New Collection
  Set pTable = pWS.OpenTable("For_ArcGIS$")
  Dim lngFamilyIndex As Long
  Dim lngTypeIndex As Long
  Dim lngGenusIndex As Long
  Dim lngSpeciesIndex As Long
  Dim lngCommonIndex As Long
  Dim lngPerAnnIndex As Long
  Dim lngC3C4Index As Long
  Dim lngNativeIndex As Long
  Dim lngGrowthFormIndex As Long
  Dim lngSPCodeIndex As Long

  lngFamilyIndex = pTable.FindField("Family")
  lngTypeIndex = pTable.FindField("Cover_Density")
  lngGenusIndex = pTable.FindField("Genus")
  lngSpeciesIndex = pTable.FindField("Species")
  lngCommonIndex = pTable.FindField("Common_Name")
  lngPerAnnIndex = pTable.FindField("perennial_annual")
  lngC3C4Index = pTable.FindField("C3_C4")
  lngNativeIndex = pTable.FindField("Native_Nonnative")
  lngGrowthFormIndex = pTable.FindField("Growth_Form")
  lngSPCodeIndex = pTable.FindField("Sp_code")

  Dim strFamily As String
  Dim strGenus As String
  Dim strJustSpecies As String
  Dim strCommon As String
  Dim strPerAnn As String
  Dim strC3C4 As String
  Dim strNative As String
  Dim strGrowthForm As String
  Dim strSPCode As String
  Dim varData As Variant
  Set pSpeciesData = New Collection

  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    If Not IsNull(pRow.Value(lngGenusIndex)) And Not IsNull(pRow.Value(lngSpeciesIndex)) And Not _
        IsNull(pRow.Value(lngFamilyIndex)) And Not IsNull(pRow.Value(lngSPCodeIndex)) Then

      strFamily = Trim(pRow.Value(lngFamilyIndex))
      strGenus = Trim(pRow.Value(lngGenusIndex))
      strJustSpecies = Trim(pRow.Value(lngSpeciesIndex))
      strSpecies = Trim(strGenus & " " & strJustSpecies)
      strType = Trim(pRow.Value(lngTypeIndex))
      strCommon = Trim(ReturnStringValFromRow(pRow, lngCommonIndex))
      strPerAnn = Trim(pRow.Value(lngPerAnnIndex))
      strC3C4 = Trim(pRow.Value(lngC3C4Index))
      strNative = Trim(ReturnStringValFromRow(pRow, lngNativeIndex))
      strGrowthForm = Trim(pRow.Value(lngGrowthFormIndex))
      strSPCode = Trim(pRow.Value(lngSPCodeIndex))
      pSpeciesData.Add Array(strFamily, strGenus, strJustSpecies, strType, strCommon, strPerAnn, strC3C4, strNative, strGrowthForm, strSPCode), strSpecies

      pReturn.Add strType, strSpecies
    Else
      Exit Do
    End If
    Set pRow = pCursor.NextRow
  Loop

  pSpeciesData.Add Array("", "", "", "Cover", "", "", "", "", "", ""), "No Cover Species Observed"
  pSpeciesData.Add Array("", "", "", "Density", "", "", "", "", "", ""), "No Density Species Observed"
  pSpeciesData.Add Array("", "", "", "Density", "Unknown Forb", "", "", "", "Forb", "UNKFOR"), "UNKFOR"

  InsertItemIntoVegCollection "Agave parryi", pSpeciesData, 8, "Succulent"
  InsertItemIntoVegCollection "Astragalus sp.", pSpeciesData, 9, "ASTR SP"
  InsertItemIntoVegCollection "Astragalus sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Asclepias sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Bothriochloa laguroides", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Bromus rubens", pSpeciesData, 7, "Non-native"
  InsertItemIntoVegCollection "Datura wrightii", pSpeciesData, 8, "Forb"
  InsertItemIntoVegCollection "Heliotropium sp.", pSpeciesData, 9, "HELI SP"
  InsertItemIntoVegCollection "Ipomoea sp.", pSpeciesData, 9, "IPOM SP"
  InsertItemIntoVegCollection "Ipomoea sp.", pSpeciesData, 8, "Forb"
  InsertItemIntoVegCollection "Menodora scabra", pSpeciesData, 8, "Subshrub"
  InsertItemIntoVegCollection "Mollugo cerviana", pSpeciesData, 7, "Non-native"
  InsertItemIntoVegCollection "Nassella leucotricha", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Muhlenbergia tenuifolia", pSpeciesData, 9, "MUHTEN"
  InsertItemIntoVegCollection "Nolina microcarpa", pSpeciesData, 8, "Succulent"
  InsertItemIntoVegCollection "Pellaea truncata", pSpeciesData, 8, "Forb"
  InsertItemIntoVegCollection "Phemeranthus aurantiacus", pSpeciesData, 8, "Forb"
  InsertItemIntoVegCollection "Phemeranthus aurantiacus", pSpeciesData, 6, "C4"
  InsertItemIntoVegCollection "Physalis sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Physalis sp.", pSpeciesData, 9, "PHYS SP"
  InsertItemIntoVegCollection "Polygala sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Polygala sp.", pSpeciesData, 9, "POLY SP"
  InsertItemIntoVegCollection "Portulaca sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "Quercus turbinella", pSpeciesData, 8, "Shrub"
  InsertItemIntoVegCollection "Symphyotrichum sp.", pSpeciesData, 7, "Native"
  InsertItemIntoVegCollection "UNKFOR", pSpeciesData, 6, "Both"
  InsertItemIntoVegCollection "UNKFOR", pSpeciesData, 7, "Both"
  InsertItemIntoVegCollection "UNKFOR", pSpeciesData, 5, "Both"

  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")

  Set ReturnSpeciesTypeColl = pReturn

End Function

Public Sub InsertItemIntoVegCollection(strSpecies As String, pSpeciesData As Collection, lngIndex As Long, strValue As String)

  Dim varVegData() As Variant
  If Not MyGeneralOperations.CheckCollectionForKey(pSpeciesData, strSpecies) Then
    MsgBox "Didn't find " & strSpecies & "..."
  Else
    varVegData = pSpeciesData.Item(strSpecies)
    pSpeciesData.Remove strSpecies
    varVegData(lngIndex) = strValue
    pSpeciesData.Add varVegData, strSpecies
  End If

End Sub

Public Sub GenerateOverstoryData_SA()

  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Dim strFinalFolder As String
  Call DeclareWorkspaces(strCombinePath, strModifiedRoot, , , , strContainerFolder, , strFinalFolder)

  Dim strNewAncillaryFolder As String
  strNewAncillaryFolder = strFinalFolder & "\Data\Ancillary_Data_CSVs"

  MyGeneralOperations.CreateNestedFoldersByPath strNewAncillaryFolder
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strNewAncillaryFolder & "\SA_5x5m_Shrub_Data_and_Quadrat_Locations.csv")

  Dim strLocationInfoPath As String
  strLocationInfoPath = MyGeneralOperations.MakeUniquedBASEName(strNewAncillaryFolder & "\Quadrat_Location_Info_Table.csv")

  Dim pFClass As IFeatureClass
  Dim lngPlotIDIndex As Long
  Dim lngSpeciesIndex As Long
  Dim lngHeightIndex As Long
  Dim lngDateIndex As Long

  Dim pDates As New Collection

  Dim strPlots() As String
  Dim pPlotColl As Collection
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strPlot As String
  Dim strSpecies As String
  Dim strAllSpecies() As String
  Dim dblHeight As Double
  Dim dblHeights() As Double
  Dim pSpeciesHeights As Collection
  Dim varPlotData() As Variant
  Dim lngMaxCount As Long
  Dim lngSpeciesCount As Long
  Dim pPointColl As New Collection
  Dim lngYear As Long
  Dim datDate As Date
  Dim pPoint As IPoint
  Dim pSpRef As IProjectedCoordinateSystem

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory

  Dim lngNSIndex As Long
  Dim lngEWIndex As Long
  Dim lngCanDiamIndex As Long
  Dim lngXIndex As Long
  Dim lngYIndex As Long
  Dim strCoord As String
  Dim strSite As String
  Dim lngSiteIndex As Long

  Dim dblLong As Double
  Dim dblLat As Double
  Dim dblNS As Double
  Dim dblEW As Double
  Dim dblCanDiam As Double
  Dim varVal As Variant

  Dim pTable As ITable
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim dblCanopies() As Double
  Dim pSpeciesCanopies As Collection

  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim varSites() As Variant
  Dim varSiteSpecifics() As Variant

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"
  Dim varArray() As Variant

  Set pQuadratColl = FillQuadratNameColl_Rev_SA(strQuadratNames, , , varSites, varSiteSpecifics, booRestrictToSite, strSiteToRestrict)

  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Summary_Data_from_JSJ\Response_from_Wade\Overstory_data_2018_All_Plots_wg.xlsx", 0)

  Set pPlotColl = New Collection
  strPlot = "6"
  strSpecies = "DASWHE"
  lngYear = 2018
  Erase strAllSpecies
  Set pSpeciesHeights = New Collection
  Set pSpeciesCanopies = New Collection

  MyGeneralOperations.AddValueToStringArray strSpecies, strAllSpecies, True
  pDates.Add lngYear, strPlot

  lngSpeciesCount = UBound(strAllSpecies) + 1
  If lngSpeciesCount > lngMaxCount Then lngMaxCount = lngSpeciesCount
  dblHeight = 0.7
  dblCanDiam = -999

  If MyGeneralOperations.CheckCollectionForKey(pSpeciesHeights, strSpecies) Then
    dblHeights = pSpeciesHeights.Item(strSpecies)
    pSpeciesHeights.Remove strSpecies
  Else
    Erase dblHeights
  End If
  If dblHeight <> -999 Then
    MyGeneralOperations.AddValueToDoubleArray dblHeight, dblHeights
    pSpeciesHeights.Add dblHeights, strSpecies
  End If

  If MyGeneralOperations.CheckCollectionForKey(pSpeciesCanopies, strSpecies) Then
    dblCanopies = pSpeciesCanopies.Item(strSpecies)
    pSpeciesCanopies.Remove strSpecies
  Else
    Erase dblCanopies
  End If
  If dblCanDiam <> -999 Then
    MyGeneralOperations.AddValueToDoubleArray dblCanDiam, dblCanopies
    pSpeciesCanopies.Add dblCanopies, strSpecies
  End If
  Erase varPlotData
  varPlotData = Array(strAllSpecies, pSpeciesHeights, pSpeciesCanopies)
  pPlotColl.Add varPlotData, strPlot

  Set pTable = pWS.OpenTable("For_ArcGIS$")
  lngSiteIndex = pTable.FindField("Site")
  lngPlotIDIndex = pTable.FindField("Plot")
  lngSpeciesIndex = pTable.FindField("Species")
  lngHeightIndex = pTable.FindField("Total_Height_M")
  lngNSIndex = pTable.FindField("NS_Transect_M")
  lngEWIndex = pTable.FindField("EW_Transect_M")
  lngCanDiamIndex = pTable.FindField("Canopy_Diameter")
  lngXIndex = pTable.FindField("X_Coord")
  lngYIndex = pTable.FindField("Y_Coord")
  lngDateIndex = pTable.FindField("Date")
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pGeoSpRef As IGeographicCoordinateSystem
  Set pGeoSpRef = MyGeneralOperations.CreateSpatialReferenceNAD83

  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing

    If Not IsNull(pRow.Value(lngSiteIndex)) Then
      strSite = pRow.Value(lngSiteIndex)
      If InStr(1, strSite, "Drainages", vbTextCompare) > 0 Then

        varVal = pRow.Value(lngPlotIDIndex)
        If IsNumeric(varVal) Then
          strPlot = Format(varVal, "0")

          If strPlot = "4" Then
            DoEvents
          End If

          varArray = pQuadratColl.Item(strPlot)
          dblLong = varArray(8)
          dblLat = varArray(9)

          datDate = pRow.Value(lngDateIndex)
          lngYear = Year(datDate)

          If Not MyGeneralOperations.CheckCollectionForKey(pDates, strPlot) Then
            pDates.Add lngYear, strPlot
          End If

          Set pPoint = New Point
          pPoint.PutCoords dblLong, dblLat
          Set pPoint.SpatialReference = pGeoSpRef
          If Not MyGeneralOperations.CheckCollectionForKey(pPointColl, strPlot) Then pPointColl.Add pPoint, strPlot

          varVal = pRow.Value(lngSpeciesIndex)
          If Not IsNull(varVal) Then
            strSpecies = Trim(CStr(varVal))
            If strSpecies <> "" And strSpecies <> "NO SP." Then

              If strSpecies = "ERICA SP" Then strSpecies = "ERIC SP"
              If strSpecies = "NOLMAC" Then strSpecies = "NOLMIC"
              If strSpecies = "AGASPP" Then strSpecies = "AGAV SP"
              If strSpecies = "OPUNSP" Then strSpecies = "OPUN SP"
              If strSpecies = "OPUSPP" Then strSpecies = "OPUN SP"
              If strSpecies = "ARCPUN" Then strSpecies = "ARCT SP"
              If strSpecies = "AGASPP" Then strSpecies = "AGAV SP"

              If strPlot = "15" Then
                DoEvents
              End If
              If strSpecies = "NOLMIC" And strPlot = "15" Then
                DoEvents
              End If

              If MyGeneralOperations.CheckCollectionForKey(pPlotColl, strPlot) Then
                varPlotData = pPlotColl.Item(strPlot)
                strAllSpecies = varPlotData(0)
                Set pSpeciesHeights = varPlotData(1)
                Set pSpeciesCanopies = varPlotData(2)
                pPlotColl.Remove strPlot
              Else
                Erase strAllSpecies
                Set pSpeciesHeights = New Collection
                Set pSpeciesCanopies = New Collection

              End If

              MyGeneralOperations.AddValueToStringArray strSpecies, strAllSpecies, True

              lngSpeciesCount = UBound(strAllSpecies) + 1
              If lngSpeciesCount > lngMaxCount Then lngMaxCount = lngSpeciesCount

              If IsNull(pRow.Value(lngHeightIndex)) Then
                dblHeight = -999
              Else
                dblHeight = pRow.Value(lngHeightIndex)

                If MyGeneralOperations.CheckCollectionForKey(pSpeciesHeights, strSpecies) Then
                  dblHeights = pSpeciesHeights.Item(strSpecies)
                  pSpeciesHeights.Remove strSpecies
                Else
                  Erase dblHeights
                End If

                MyGeneralOperations.AddValueToDoubleArray dblHeight, dblHeights
                pSpeciesHeights.Add dblHeights, strSpecies
              End If

              If IsNull(pRow.Value(lngHeightIndex)) Then
                dblCanDiam = -999
              Else
                dblCanDiam = pRow.Value(lngCanDiamIndex)

                If MyGeneralOperations.CheckCollectionForKey(pSpeciesCanopies, strSpecies) Then
                  dblCanopies = pSpeciesCanopies.Item(strSpecies)
                  pSpeciesCanopies.Remove strSpecies
                Else
                  Erase dblCanopies
                End If

                MyGeneralOperations.AddValueToDoubleArray dblCanDiam, dblCanopies
                pSpeciesCanopies.Add dblCanopies, strSpecies
              End If
            End If

            Erase varPlotData
            varPlotData = Array(strAllSpecies, pSpeciesHeights, pSpeciesCanopies)
            pPlotColl.Add varPlotData, strPlot
          End If
        End If
      End If
    End If

    Set pRow = pCursor.NextRow
  Loop

  If Not MyGeneralOperations.CheckCollectionForKey(pDates, "1") Then pDates.Add 2018, "1"
  If Not MyGeneralOperations.CheckCollectionForKey(pDates, "21") Then pDates.Add 2018, "21"

  Dim strOutput As String
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim dblMean As Double
  Dim dblMin As Double
  Dim dblMax As Double
  Dim lngCount As Long
  Dim strSpeciesStats As String
  Dim strSpeciesCanStats As String
  Dim strParentMaterial As String
  Dim dblElev As Double
  Dim dblAspect As Double
  Dim dblSlope As Double
  Dim dblUTME As Double
  Dim dblUTMN As Double
  Dim strYear As String
  Dim dblYearMeasured As Double
  Dim strComment As String
  Dim strDrainage As String
  Dim pEnvironmentColl As Collection
  Dim varData() As Variant
  Dim dblLongitude As Double
  Dim dblLatitude As Double

  Dim strOutputLocationInfo As String

  strOutput = """Site"",""Quadrat"",""Drainage"",""Num_Species"","
  For lngIndex = 1 To lngMaxCount
    strOutput = strOutput & """Species_" & Format(lngIndex, "0") & """,""S" & Format(lngIndex, "0") & "_Ht_n_min_mean_max"",""S" & _
        Format(lngIndex, "0") & "_CanDiam_n_min_mean_max"","
  Next lngIndex
  strOutput = strOutput & """Parent_Material"",""Elevation"",""Aspect"",""Slope_Percent"",""Easting_NAD_1983_UTM_12""," & _
      """Northing_NAD_1983_UTM_12"",""Longitude_NAD_1983"",""Latitude_NAD_1983"",""Year_Measured""" & vbCrLf

  strOutputLocationInfo = """Site"",""Quadrat"",""Drainage"",""Parent_Material"",""Elevation"",""Aspect"",""Slope_Percent"",""Easting_NAD_1983_UTM_12""," & _
      """Northing_NAD_1983_UTM_12"",""Longitude_NAD_1983"",""Latitude_NAD_1983"",""Year_Measured""" & vbCrLf

  For lngIndex = 1 To 24
    DoEvents
    strPlot = Format(lngIndex, "0")
    If strPlot = "15" Then
      DoEvents
    End If
    Set pEnvironmentColl = ReturnSAEnvironmentalData
    varData = pEnvironmentColl.Item(strPlot)
    strDrainage = varData(0)
    strParentMaterial = varData(1)
    dblElev = varData(2)
    dblAspect = varData(3)
    dblSlope = varData(4)
    dblUTME = varData(5)
    dblUTMN = varData(6)
    dblLongitude = varData(7)
    dblLatitude = varData(8)

    If MyGeneralOperations.CheckCollectionForKey(pDates, strPlot) Then
      lngYear = pDates.Item(strPlot)
      strYear = Format(lngYear, "0")
    Else
      strYear = ""
    End If

    If MyGeneralOperations.CheckCollectionForKey(pPlotColl, strPlot) Then
      varPlotData = pPlotColl.Item(strPlot)
      strAllSpecies = varPlotData(0)
      Set pSpeciesHeights = varPlotData(1)
      Set pSpeciesCanopies = varPlotData(2)
      lngSpeciesCount = UBound(strAllSpecies) + 1
      Debug.Print "Plot " & strPlot & " has " & Format(lngSpeciesCount, "0") & " species"

      QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)

      strOutput = strOutput & "Natural Drainages," & strPlot & "," & strDrainage & "," & Format(lngSpeciesCount, "0") & ","
      strOutputLocationInfo = strOutputLocationInfo & "Natural Drainages," & strPlot & "," & strDrainage & ","

      For lngIndex2 = 0 To lngMaxCount - 1
        If UBound(strAllSpecies) >= lngIndex2 Then
          strSpecies = strAllSpecies(lngIndex2)
          dblHeights = pSpeciesHeights.Item(strSpecies)
          Call MyGeneralOperations.BasicStatsFromArraySimpleFast(dblHeights, False, , lngCount, dblMin, dblMax, dblMean)
          strSpeciesStats = Format(lngCount, "0") & "/" & Format(dblMin, "0.0") & "/" & Format(dblMean, "0.0") & "/" & Format(dblMax, "0.0")
          If MyGeneralOperations.CheckCollectionForKey(pSpeciesCanopies, strSpecies) Then
            dblCanopies = pSpeciesCanopies.Item(strSpecies)
          Else
            Erase dblCanopies
          End If
          If IsDimmed(dblCanopies) Then
            Call MyGeneralOperations.BasicStatsFromArraySimpleFast(dblCanopies, False, , lngCount, dblMin, dblMax, dblMean)
            strSpeciesCanStats = Format(lngCount, "0") & "/" & Format(dblMin, "0.0") & "/" & Format(dblMean, "0.0") & "/" & Format(dblMax, "0.0")
          Else
            strSpeciesCanStats = "0/NA/NA/NA"
          End If
        Else
          strSpecies = ""
          strSpeciesStats = ""
          strSpeciesCanStats = ""
        End If

        strOutput = strOutput & strSpecies & "," & strSpeciesStats & "," & strSpeciesCanStats & ","
      Next lngIndex2

      strOutput = strOutput & strParentMaterial & "," & Format(dblElev, "0.00") & "," & Format(dblAspect, "0") & "," & _
          Format(dblSlope, "0%") & "," & Format(dblUTME, "0") & "," & Format(dblUTMN, "0") & "," & Format(dblLongitude, "0.000000") & _
          "," & Format(dblLatitude, "0.000000") & "," & strYear & vbCrLf
      strOutputLocationInfo = strOutputLocationInfo & strParentMaterial & "," & Format(dblElev, "0.00") & "," & Format(dblAspect, "0") & "," & _
          Format(dblSlope, "0%") & "," & Format(dblUTME, "0") & "," & Format(dblUTMN, "0") & "," & Format(dblLongitude, "0.000000") & _
          "," & Format(dblLatitude, "0.000000") & "," & strYear & vbCrLf

    Else
      Debug.Print "Plot " & strPlot & " has 0 species"

      strOutput = strOutput & "Natural Drainages," & strPlot & "," & strDrainage & ",0,"
      strOutputLocationInfo = strOutputLocationInfo & "Natural Drainages," & strPlot & "," & strDrainage & ","

      For lngIndex2 = 0 To lngMaxCount - 1
        strSpecies = ""
        strSpeciesStats = ""
        strSpeciesCanStats = ""

        strOutput = strOutput & strSpecies & "," & strSpeciesStats & "," & strSpeciesCanStats & ","
      Next lngIndex2

      strOutput = strOutput & strParentMaterial & "," & Format(dblElev, "0.00") & "," & Format(dblAspect, "0") & "," & _
          Format(dblSlope, "0%") & "," & Format(dblUTME, "0") & "," & Format(dblUTMN, "0") & "," & Format(dblLongitude, "0.000000") & _
          "," & Format(dblLatitude, "0.000000") & "," & strYear & vbCrLf
      strOutputLocationInfo = strOutputLocationInfo & strParentMaterial & "," & Format(dblElev, "0.00") & "," & Format(dblAspect, "0") & "," & _
          Format(dblSlope, "0%") & "," & Format(dblUTME, "0") & "," & Format(dblUTMN, "0") & "," & Format(dblLongitude, "0.000000") & _
          "," & Format(dblLatitude, "0.000000") & "," & strYear & vbCrLf

    End If

  Next lngIndex

  MyGeneralOperations.WriteTextFile strExportPath, strOutput
  MyGeneralOperations.WriteTextFile strLocationInfoPath, strOutputLocationInfo

  Debug.Print "Done..."

End Sub

Public Function ReturnDDfromString(strCoord As String) As Double

  Dim dblDeg As Double
  Dim dblMin As Double
  Dim dblSec As Double
  Dim strSplit() As String
  strSplit = Split(strCoord, " ")
  dblDeg = CDbl(strSplit(0))
  dblMin = CDbl(strSplit(1))
  dblSec = CDbl(strSplit(2))

  Dim booIsNeg As Double
  booIsNeg = dblDeg < 0

  dblDeg = Abs(dblDeg)

  Dim dblReturn As Double
  dblReturn = dblDeg + (dblMin / 60) + (dblSec / 3600)
  If booIsNeg Then dblReturn = dblReturn * -1
  ReturnDDfromString = dblReturn
End Function

Public Function ReturnSAEnvironmentalData() As Collection

  Dim pReturn As New Collection
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pWS As IFeatureWorkspace
  Dim pRastWS As IRasterWorkspaceEx
  Dim pWSFact As IWorkspaceFactory

  Dim pDrainageA As IPolygon
  Dim pDrainageB As IPolygon
  Dim pDrainageC As IPolygon
  Dim pDrainageD As IPolygon

  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\ND & Plot 3 Maps & Shapefiles", 0)
  Set pDrainageA = ReturnFirstFeature(pWS.OpenFeatureClass("Drainage A"))
  Set pDrainageB = ReturnFirstFeature(pWS.OpenFeatureClass("Drainage B"))
  Set pDrainageC = ReturnFirstFeature(pWS.OpenFeatureClass("Drainage C"))
  Set pDrainageD = ReturnFirstFeature(pWS.OpenFeatureClass("Drainage D"))

  Dim pDEM As IRaster
  Dim pAspect As IRaster
  Dim pSlope As IRaster
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pRastWS = pWSFact.OpenFromFile("D:\GIS_Data\DEM_Stuff\Full_DEM_Data.gdb", 0)
  Set pDEM = pRastWS.OpenRasterDataset("All_US_NoNull").CreateDefaultRaster
  Set pAspect = pRastWS.OpenRasterDataset("All_US_Aspect").CreateDefaultRaster
  Set pSlope = pRastWS.OpenRasterDataset("All_US_Slope").CreateDefaultRaster

  Dim pPointsFClass As IFeatureClass
  Dim varData() As Variant
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPoint As IPoint
  Dim pDrainFCursor As IFeatureCursor
  Dim pDrainFeature As IFeature
  Dim pSpFilt As ISpatialFilter

  Dim strDrainage As String
  Dim dblElev As Double
  Dim dblAspect As Double
  Dim dblSlope As Double
  Dim dblUTME As Double
  Dim dblUTMN As Double
  Dim lngYearMeasured As Long
  Dim strComment As String
  Dim strParentMaterial As String
  Dim strPlot As String
  Dim lngPlotIndex As Long
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  Dim pNAD83 As ISpatialReference
  Dim pClone As IClone
  Dim pGeoPoint As IPoint

  Set pNAD83 = MyGeneralOperations.CreateSpatialReferenceNAD83

  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Reference_Datasets.gdb", 0)
  Set pPointsFClass = pWS.OpenFeatureClass("Sierra_Ancha_Quadrat_Locations_UTM12_NAD83")
  lngPlotIndex = pPointsFClass.FindField("Quadrat")

  Dim pParentMaterialColl As Collection
  Set pParentMaterialColl = ReturnParentMaterialPerQuadrat

  Dim pGeoSpRef As IGeographicCoordinateSystem
  Set pGeoSpRef = MyGeneralOperations.CreateSpatialReferenceNAD83

  Set pFCursor = pPointsFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    DoEvents
    strPlot = Format(pFeature.Value(lngPlotIndex), "0")

    Set pPoint = pFeature.ShapeCopy
    dblUTME = pPoint.x
    dblUTMN = pPoint.Y
    Set pClone = pPoint
    Set pGeoPoint = pClone.Clone
    pGeoPoint.Project pNAD83
    dblLongitude = pGeoPoint.x
    dblLatitude = pGeoPoint.Y

    strDrainage = ReturnIntersectingDrainage(pPoint, pDrainageA, pDrainageB, pDrainageC, pDrainageD)

    pPoint.Project pGeoSpRef

    dblElev = GridFunctions.CellValue4CellInterp(pPoint, pDEM)
    dblAspect = GridFunctions.CellValue4CellInterp_Direction(pPoint, pAspect)
    dblSlope = MyGeometricOperations.DegToPercent(GridFunctions.CellValue4CellInterp(pPoint, pSlope))

    If MyGeneralOperations.CheckCollectionForKey(pParentMaterialColl, strPlot) Then
      strParentMaterial = pParentMaterialColl.Item(strPlot)
    Else
      strParentMaterial = ""
    End If

    varData = Array(strDrainage, strParentMaterial, dblElev, dblAspect, dblSlope, dblUTME, dblUTMN, dblLongitude, dblLatitude)
    pReturn.Add varData, strPlot

    Set pFeature = pFCursor.NextFeature
  Loop

  Set ReturnSAEnvironmentalData = pReturn

End Function

Public Function ReturnParentMaterialPerQuadrat() As Collection

  Dim pPMColl As New Collection
  pPMColl.Add "Q", "1"
  pPMColl.Add "Q", "2"
  pPMColl.Add "Q", "3"
  pPMColl.Add "Q", "4"
  pPMColl.Add "Q", "5"
  pPMColl.Add "Q", "6"
  pPMColl.Add "Q", "7"
  pPMColl.Add "Q", "8"
  pPMColl.Add "Q", "9"
  pPMColl.Add "Q", "10"
  pPMColl.Add "Q", "11"
  pPMColl.Add "Q", "12"
  pPMColl.Add "D", "13"
  pPMColl.Add "D", "14"
  pPMColl.Add "D", "15"
  pPMColl.Add "Q", "16"
  pPMColl.Add "Q", "17"
  pPMColl.Add "D", "18"
  pPMColl.Add "D", "19"
  pPMColl.Add "D", "20"
  pPMColl.Add "Q", "21"
  pPMColl.Add "Q", "22"
  pPMColl.Add "Q", "23"
  pPMColl.Add "Q", "24"
  Set ReturnParentMaterialPerQuadrat = pPMColl

End Function

Public Function ReturnIntersectingDrainage(pPoint As IPoint, pDrainageA As IPolygon, pDrainageB As IPolygon, _
    pDrainageC As IPolygon, pDrainageD As IPolygon) As String

  Dim pTransform As IGeoTransformation
  Dim pClone As IClone
  Dim pNewPoint As IPoint
  Dim pGeom As IGeometry2
  Dim pRelOp As IRelationalOperator

  Set pClone = pPoint
  Set pNewPoint = pClone.Clone
  Set pGeom = pNewPoint

  Set pTransform = MyGeneralOperations.CreateNAD83_WGS84_GeoTransformationFlagstaff
  pGeom.ProjectEx pDrainageA.SpatialReference, esriTransformForward, pTransform, False, 0, 0

  Set pRelOp = pGeom
  If Not pRelOp.Disjoint(pDrainageA) Then
    ReturnIntersectingDrainage = "A"
  ElseIf Not pRelOp.Disjoint(pDrainageB) Then
    ReturnIntersectingDrainage = "B"
  ElseIf Not pRelOp.Disjoint(pDrainageC) Then
    ReturnIntersectingDrainage = "C"
  ElseIf Not pRelOp.Disjoint(pDrainageD) Then
    ReturnIntersectingDrainage = "D"
  Else
    ReturnIntersectingDrainage = ""
  End If

  Set pTransform = Nothing
  Set pClone = Nothing
  Set pNewPoint = Nothing
  Set pGeom = Nothing
  Set pRelOp = Nothing

End Function

Public Function ReturnFirstFeature(pFClass As IFeatureClass) As IGeometry

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Set pFCursor = pFClass.Search(Nothing, True)
  Set pFeature = pFCursor.NextFeature
  Set ReturnFirstFeature = pFeature.ShapeCopy
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Function

Public Sub OrganizeData_SA()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Dim pRedigitizeColl As Collection
  Set pRedigitizeColl = ReturnReplacementColl_SA
  Dim pDataObj As New MSForms.DataObject

  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Dim lngCount As Long
  Dim lngIndex As Long
  Dim strPath As String
  Dim strModPath As String
  Dim lngCounter As Long

  Dim strQuadrat As String
  Dim strReplaceName As String

  Dim strExt As String
  Dim booTransfer As Boolean
  Dim strFilename As String
  Dim strNewDir As String

  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim strCombinePath As String
  Dim strSetFolder As String
  Call DeclareWorkspaces(strCombinePath, , , , , strSetFolder)

  If Not aml_func_mod.ExistFileDir(strCombinePath) Then
    MyGeneralOperations.CreateNestedFoldersByPath strCombinePath
  End If
  If Not aml_func_mod.ExistFileDir(strSetFolder & "\Description_of_Analysis.docx") Then
    CopyFile "D:\arcGIS_stuff\consultation\Margaret_Moore\Data_to_include_in_publication\Description_of_Analysis.docx", _
      strSetFolder & "\Description_of_Analysis.docx", 0
  End If

  Dim strDir As String
  Dim pAllPaths As esriSystem.IStringArray
  Dim varCheckArray() As Variant
  Dim strCheckPathReport
  Dim pDataset As IDataset

  Dim strSourcePath1 As String

  Dim strAllPaths() As String
  strAllPaths = ReturnInitialDataUpTo2020

  varCheckArray = BuildCheckArray_SA(strAllPaths)

  Dim pConvertNamesOldTo2020 As Collection
  Dim pConvertNames2020ToOld As Collection
  Dim varNameLinks() As Variant
  Call FillNameConverters_SA_Original(varNameLinks, pConvertNames2020ToOld, pConvertNamesOldTo2020)

  lngCount = UBound(strAllPaths) + 1
  lngCounter = 0
  Debug.Print "Round 1: " & Format(lngCount, "#,##0") & " paths found..."
  Dim pCopyFClass As IFeatureClass
  Dim pDoneColl As New Collection
  Dim pUnknownSpRef As IUnknownCoordinateSystem
  Set pUnknownSpRef = New UnknownCoordinateSystem
  Dim pGeoDataset As IGeoDataset
  Dim lngYear As Long

  Dim pYearSpecialConversions As New Collection
  pYearSpecialConversions.Add "1935", "1395"

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To UBound(strAllPaths)
      pProg.Step

      strPath = strAllPaths(lngIndex)
      strSourcePath1 = aml_func_mod.ReturnDir3(strPath, False)

      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        If InStr(1, strPath, "VBA", vbTextCompare) > 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)

        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          lngYear = SierraAncha.ReturnYearFromName(strFilename)

          If MyGeneralOperations.CheckCollectionForKey(pYearSpecialConversions, Format(lngYear, "0")) Then
            lngYear = CLng(pYearSpecialConversions.Item(Format(lngYear, "0")))
          End If

          If strFilename = "Q13_1955_DCopy" Then
            Debug.Print "...Skipping '" & strFilename & "'..."
          Else

            If MyGeneralOperations.CheckCollectionForKey(pRedigitizeColl, strFilename) Then

              If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strFilename) Then
                Set pDataset = pRedigitizeColl.Item(strFilename)

                strQuadrat = pConvertNames2020ToOld.Item(strFilename)

                If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                  strFilename = Replace(strFilename, "_CF", "_C", , , vbTextCompare)
                  strFilename = Replace(strFilename, "_DF", "_D", , , vbTextCompare)
                  strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_" & Format(lngYear, "0") & _
                      IIf(StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0, "_C", "_D") & "." & strExt

                  Set pCopyFClass = CopyFeatureClassToShapefile(pDataset, strModPath)
                  Debug.Print "...Using redigitized feature class '" & pDataset.BrowseName & "..."
                  pDoneColl.Add True, strFilename

                  UpdateCheckArray varCheckArray, strPath
                End If ' END JUST DOING NATURAL DRAINAGES
              Else
                Debug.Print "...Already copied over '" & strFilename & "..."
              End If

            Else

              If lngYear > 1955 Then
                DoEvents
              End If

              If StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0 Or _
                  StrComp(Right(strFilename, 2), "_D", vbTextCompare) = 0 Or _
                  StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
                  StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Then

                UpdateCheckArray varCheckArray, strPath

                If MyGeneralOperations.CheckCollectionForKey(pConvertNames2020ToOld, strFilename) Then
                  strQuadrat = pConvertNames2020ToOld.Item(strFilename)

                  If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                    strFilename = Replace(strFilename, "_CF", "_C", , , vbTextCompare)
                    strFilename = Replace(strFilename, "_DF", "_D", , , vbTextCompare)
                    strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_" & Format(lngYear, "0") & _
                        IIf(StrComp(Right(strFilename, 2), "_C", vbTextCompare) = 0, "_C", "_D") & "." & strExt

                    If Not aml_func_mod.ExistFileDir(strModPath) Then
                      strDir = aml_func_mod.ReturnDir3(strModPath, False)
                      If Not aml_func_mod.ExistFileDir(strDir) Then _
                        MyGeneralOperations.CreateNestedFoldersByPath strDir
                        lngCounter = lngCounter + 1
                        CopyFile strPath, strModPath, True
                    End If
                  End If ' END JUST DOING NATURAL DRAINAGES
                Else
                  Debug.Print "Failed to find '" & strFilename & "'" & vbCrLf & _
                      "...Path = '" & strPath & "'..."
                End If
              End If
            End If
          End If
        End If

      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then

      pDataObj.SetText strCheckPathReport
      pDataObj.PutInClipboard
      Set pDataObj = Nothing
      Debug.Print "Original Data: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  DoEvents

  Dim strSourcePath8 As String
  strSourcePath8 = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\2021\Final"

  Dim pConvertNamesOldTo2021 As Collection
  Dim pConvertNames2021ToOld As Collection
  Dim varNameLinks_2021() As Variant
  Call FillNameConverters_2021(varNameLinks_2021, pConvertNames2021ToOld, pConvertNamesOldTo2021)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath8, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 2: " & Format(lngCount, "#,##0") & " paths found..."

  Dim booRunTransform As Boolean ' <--------  SPECIAL CASE TO TRANSFORM DATA IN 10 X 10 SPACE

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      booRunTransform = False ' <--------  SPECIAL CASE TO TRANSFORM DATA IN 10 X 10 SPACE
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)

          If strFilename = "Q_ND_1_2021_DF" Then
            DoEvents
          End If

          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2021ToOld, strFilename) Then
              strQuadrat = pConvertNames2021ToOld.Item(strFilename)

              If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2021" & _
                    IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                        (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                        (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
                strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

                If Not aml_func_mod.ExistFileDir(strModPath) Then
                  strDir = aml_func_mod.ReturnDir3(strModPath, False)
                  If Not aml_func_mod.ExistFileDir(strDir) Then _
                    MyGeneralOperations.CreateNestedFoldersByPath strDir
                    lngCounter = lngCounter + 1
                    CopyFile strPath, strModPath, True
                End If
              End If ' END JUST DOING NATURAL DRAINAGES
            Else
              Debug.Print "Failed to find '" & strFilename & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2021: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim strSourcePath9 As String
  strSourcePath9 = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\2022\Final"

  Dim pConvertNamesOldTo2022 As Collection
  Dim pConvertNames2022ToOld As Collection
  Dim varNameLinks_2022() As Variant
  Call FillNameConverters_2022(varNameLinks_2022, pConvertNames2022ToOld, pConvertNamesOldTo2022)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath9, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 3 [2022]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          If strFilename = "Q_FV_22241_2022_D_NO_SpF" Then strFilename = "Q_FV_22241_2022_DF" ' FOR NOV. 2022 DATA
          If strFilename = "Q_FV_22241_2022_C_NO_SpF" Then strFilename = "Q_FV_22241_2022_CF" ' FOR NOV. 2022 DATA
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR NOV. 2022 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Sp", "", , , vbTextCompare) ' FOR NOV. 2022 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2022ToOld, strFilename) Then
              strQuadrat = pConvertNames2022ToOld.Item(strFilename)

              If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2022" & _
                    IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                        (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                        (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
                strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

                If Not aml_func_mod.ExistFileDir(strModPath) Then
                  strDir = aml_func_mod.ReturnDir3(strModPath, False)
                  If Not aml_func_mod.ExistFileDir(strDir) Then _
                    MyGeneralOperations.CreateNestedFoldersByPath strDir
                    lngCounter = lngCounter + 1
                    CopyFile strPath, strModPath, True
                End If
              End If ' END JUST DOING NATURAL DRAINAGES
            Else
              Debug.Print "Failed to find '" & strFilename & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2022: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim strSourcePath10 As String
  strSourcePath10 = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\2023\Final"

  Dim pConvertNamesOldTo2023 As Collection
  Dim pConvertNames2023ToOld As Collection
  Dim varNameLinks_2023() As Variant
  Call FillNameConverters_2023(varNameLinks_2023, pConvertNames2023ToOld, pConvertNamesOldTo2023)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath10, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 4 [2023]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR NOV. 2023 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Sp", "", , , vbTextCompare) ' FOR NOV. 2023 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2023ToOld, strFilename) Then
              strQuadrat = pConvertNames2023ToOld.Item(strFilename)

              If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2023" & _
                    IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                        (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                        (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
                strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

                If Not aml_func_mod.ExistFileDir(strModPath) Then
                  strDir = aml_func_mod.ReturnDir3(strModPath, False)
                  If Not aml_func_mod.ExistFileDir(strDir) Then _
                    MyGeneralOperations.CreateNestedFoldersByPath strDir
                    lngCounter = lngCounter + 1
                    CopyFile strPath, strModPath, True
                End If
              End If ' END JUST DOING NATURAL DRAINAGES
            Else
              Debug.Print "Failed to find '" & strFilename & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2023: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim strSourcePath11 As String
  strSourcePath11 = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\2024\Final"

  Dim pConvertNamesOldTo2024 As Collection
  Dim pConvertNames2024ToOld As Collection
  Dim varNameLinks_2024() As Variant
  Call FillNameConverters_2024(varNameLinks_2024, pConvertNames2024ToOld, pConvertNamesOldTo2024)
  Set pAllPaths = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourcePath11, "")
  varCheckArray = BuildCheckArray(pAllPaths)

  lngCount = pAllPaths.Count
  lngCounter = 0
  Debug.Print "Round 5 [2024]: " & Format(lngCount, "#,##0") & " paths found..."

  If lngCount > 0 Then

    pSBar.ShowProgressBar "Copying Files...", 0, lngCount, 1, True
    pProg.position = 0

    For lngIndex = 0 To lngCount - 1
      If lngIndex Mod 500 = 0 Then DoEvents
      strPath = pAllPaths.Element(lngIndex)
      pProg.Step
      If StrComp(Right(strPath, 5), ".lock", vbTextCompare) <> 0 Then
        If lngIndex Mod 100 = 0 Then
          DoEvents
        End If
        strExt = aml_func_mod.GetExtensionText(strPath)
        booTransfer = False

        If StrComp(strExt, "cpg", vbTextCompare) = 0 Or StrComp(strExt, "dbf", vbTextCompare) = 0 Or _
            StrComp(strExt, "sbn", vbTextCompare) = 0 Or StrComp(strExt, "sbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "shp", vbTextCompare) = 0 Or StrComp(strExt, "shx", vbTextCompare) = 0 Or _
            StrComp(strExt, "prj", vbTextCompare) = 0 Or StrComp(strExt, "fbn", vbTextCompare) = 0 Or _
            StrComp(strExt, "ain", vbTextCompare) = 0 Or StrComp(strExt, "fbx", vbTextCompare) = 0 Or _
            StrComp(strExt, "aih", vbTextCompare) = 0 Or StrComp(strExt, "ixs", vbTextCompare) = 0 Or _
            StrComp(strExt, "mxs", vbTextCompare) = 0 Or StrComp(strExt, "qix", vbTextCompare) = 0 Or _
            StrComp(strExt, "atx", vbTextCompare) = 0 Then
          booTransfer = True
        ElseIf StrComp(Right(strPath, 8), ".shp.xml", vbTextCompare) = 0 Then
          booTransfer = True
          strExt = ".shp.xml"
        End If

        If booTransfer Then
          strFilename = aml_func_mod.ReturnFilename2(strPath)
          strFilename = Replace(strFilename, ".shp.xml", "", , , vbTextCompare)
          strFilename = aml_func_mod.ClipExtension2(strFilename)
          strReplaceName = Replace(strFilename, "_CF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_DF", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_C_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_D_F", "", , , vbTextCompare)
          strReplaceName = Replace(strReplaceName, "_NO_SpF", "", , , vbTextCompare) ' FOR NOV. 2023 DATA
          strReplaceName = Replace(strReplaceName, "_NO_Sp", "", , , vbTextCompare) ' FOR NOV. 2023 DATA

          If StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0 Or _
              StrComp(Right(strFilename, 3), "_DF", vbTextCompare) = 0 Or _
              InStr(1, strFilename, "_CF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_DF_", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_C_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_D_F", vbTextCompare) > 0 Or _
              InStr(1, strFilename, "_NO_Sp", vbTextCompare) > 0 Then

            UpdateCheckArray varCheckArray, strPath

            If MyGeneralOperations.CheckCollectionForKey(pConvertNames2024ToOld, strFilename) Then
              strQuadrat = pConvertNames2024ToOld.Item(strFilename)

              If InStr(1, strQuadrat, "Natural_Drainages", vbTextCompare) > 0 Then

                strModPath = strCombinePath & "\" & strQuadrat & "\" & strQuadrat & "_2024" & _
                    IIf((InStr(1, strFilename, "_CF_", vbTextCompare) > 0) Or _
                        (StrComp(Right(strFilename, 3), "_CF", vbTextCompare) = 0) Or _
                        (StrComp(Right(strFilename, 4), "_C_F", vbTextCompare) = 0), "_C", "_D") & "." & strExt
                strModPath = Replace(strModPath, "..", ".", , , vbTextCompare)

                If Not aml_func_mod.ExistFileDir(strModPath) Then
                  strDir = aml_func_mod.ReturnDir3(strModPath, False)
                  If Not aml_func_mod.ExistFileDir(strDir) Then _
                    MyGeneralOperations.CreateNestedFoldersByPath strDir
                    lngCounter = lngCounter + 1
                    CopyFile strPath, strModPath, True
                End If
              End If ' END JUST DOING NATURAL DRAINAGES
            Else
              Debug.Print "Failed to find '" & strFilename & "'" & vbCrLf & _
                  "...Path = '" & strPath & "'..."
            End If
          End If
        Else

        End If
      End If
    Next lngIndex

    strCheckPathReport = ReturnMissingShapefiles(varCheckArray)
    If strCheckPathReport <> "" Then
      Debug.Print "2024: Check Following Shapefiles" & vbCrLf & strCheckPathReport
    End If

    pSBar.HideProgressBar
    pProg.position = 0

  End If

  Dim strQuadrats() As String
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim lngFeatCount As Long
  Dim pQuadData As Collection
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev_SA(strQuadrats, pPlotToQuadratConversion, pQuadratToPlotConversion, _
      varSites, varSitesSpecific, booRestrictToSite, strSiteToRestrict)

  Dim pSHPfiles As esriSystem.IStringArray
  Set pSHPfiles = MyGeneralOperations.ReturnFilesFromNestedFolders(strCombinePath, "shp")

  Debug.Print "pSHPfiles.Count = " & Format(pSHPfiles.Count, "0")

  Dim pDone1 As New Collection
  Dim strNames1() As String
  Dim lngNameIndex As Long
  Dim pWSFact As IWorkspaceFactory
  Dim pWS As IFeatureWorkspace

  Dim pDatasets As IEnumDataset
  Dim strName As String
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim lngSrcSpeciesNameIndex As Long
  Dim lngVerbSpeciesNameIndex As Long
  Dim lngVerbTypeIndex As Long
  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pClone As IClone
  Dim pFeature As IFeature
  Dim strPrjPath As String

  pSBar.ShowProgressBar "Adding Verbatim Fields...", 0, pSHPfiles.Count, 1, True
  pProg.position = 0
  lngNameIndex = -1
  For lngIndex = 0 To pSHPfiles.Count - 1
    strPath = pSHPfiles.Element(lngIndex)
    strPrjPath = strPath
    strPrjPath = aml_func_mod.SetExtension(strPrjPath, "prj")
    If aml_func_mod.ExistFileDir(strPrjPath) Then
      Kill strPrjPath
    End If
    pProg.Step
    If lngIndex Mod 25 = 0 Then
      DoEvents
    End If
    Debug.Print MyGeneralOperations.SpacesInFrontOfText(Format(lngIndex, "#,##0"), 5) & "] " & aml_func_mod.ReturnFilename2(strPath)
    strDir = aml_func_mod.ReturnDir3(strPath, False)
    Set pWSFact = New ShapefileWorkspaceFactory
    Set pWS = pWSFact.OpenFromFile(strDir, 0)
    Set pFClass = pWS.OpenFeatureClass(aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strPath)))
    AddVerbatimFields_SA pFClass, pQuadData

  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  Debug.Print "Done..."

ClearMemory:
  Set pRedigitizeColl = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pAllPaths = Nothing
  Erase varCheckArray
  Set pDataset = Nothing
  Erase strAllPaths
  Set pConvertNamesOldTo2020 = Nothing
  Set pConvertNames2020ToOld = Nothing
  Erase varNameLinks
  Set pCopyFClass = Nothing
  Set pDoneColl = Nothing
  Set pUnknownSpRef = Nothing
  Set pGeoDataset = Nothing
  Set pDataObj = Nothing
  Erase strQuadrats
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pQuadData = Nothing
  Erase varSites
  Erase varSitesSpecific
  Set pSHPfiles = Nothing
  Set pDone1 = Nothing
  Erase strNames1
  Set pWSFact = Nothing
  Set pWS = Nothing
  Set pDatasets = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pClone = Nothing
  Set pFeature = Nothing

End Sub

Public Function ReturnInitialDataUpTo2020() As String()

  Dim varYears() As Variant
  varYears = Array(1935, 1936, 1940, 1941, 1942, 1952, 1955, 2017, 2018, 2019, 2020)

  Dim pShapefiles As esriSystem.IStringArray

  Dim strFinalArray() As String
  Dim lngArrayIndex As Long
  Dim lngYearIndex As Long
  Dim lngYear As Long
  Dim strSourceFolder As String
  Dim lngShpIndex As Long
  Dim strShapeName As String
  Dim strClipName As String

  Dim pDoneColl As New Collection

  strSourceFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\"

  Dim strReport As String

  lngArrayIndex = -1
  Set pDoneColl = New Collection

  For lngYearIndex = 0 To UBound(varYears)
    lngYear = varYears(lngYearIndex)
    Set pShapefiles = MyGeneralOperations.ReturnFilesFromNestedFolders2(strSourceFolder & _
        Format(lngYear, "0") & "\Final", "")
    Debug.Print Format(lngYear, "0") & "............."
    For lngShpIndex = 0 To pShapefiles.Count - 1
      strShapeName = pShapefiles.Element(lngShpIndex)
      strClipName = strShapeName ' aml_func_mod.ClipExtension2(aml_func_mod.ReturnFilename2(strShapeName))

        pDoneColl.Add True, strClipName
        lngArrayIndex = lngArrayIndex + 1
        ReDim Preserve strFinalArray(lngArrayIndex)
        strFinalArray(lngArrayIndex) = strClipName
    Next lngShpIndex

  Next lngYearIndex

  ReturnInitialDataUpTo2020 = strFinalArray

End Function

Public Sub CreateFinalTables_SA()

  Debug.Print "-----------------------------------"

  GenerateOverstoryData_SA

  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFinalTable As ITable
  Dim pAddTable As ITable

  Dim strOrigRoot As String
  Dim strModifiedRoot As String
  Dim strShiftRoot As String
  Dim strFinalFolder As String
  Dim strExportBase As String
  Dim strSetFolder As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, strExportBase, strModifiedRoot, strSetFolder, , strFinalFolder)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection

  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim varSites() As Variant
  Dim varSiteSpecifics() As Variant
  Dim varArray() As Variant

  Dim pPlotToQuadratColl As Collection
  Dim pQuadratToPlotColl As Collection
  Set pQuadratColl = FillQuadratNameColl_Rev_SA(strQuadratNames, pPlotToQuadratColl, pQuadratToPlotColl, varSites, varSiteSpecifics)

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace

  Dim strFClassName As String
  Dim strNameSplit() As String
  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Dim pWStoUpdate As IWorkspace
  If aml_func_mod.ExistFileDir(strFinalFolder & "\Combined_by_Site.gdb") Then
    Set pWStoUpdate = pNewWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
  Else
    Set pWStoUpdate = pNewWSFact.OpenFromFile(strFinalFolder & "\Data\Quadrat_Spatial_Data\Combined_by_Site.gdb", 0)
  End If
  Dim pEnumDataset As IEnumDataset
  Dim pUpdateDataset As IDataset
  Dim pFClass As IFeatureClass

  Dim strNewAncillaryFolder As String
  strNewAncillaryFolder = strFinalFolder & "\Data\Ancillary_Data_CSVs"

  MyGeneralOperations.CreateNestedFoldersByPath strNewAncillaryFolder

  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strFinalFolder & "\Data\Ancillary_Data_GDB")
  Dim pWS2 As IWorkspace2
  Dim pFeatWS As IFeatureWorkspace
  Set pWS2 = pNewWS
  Set pFeatWS = pWS2

  Dim pFCursor As IFeatureCursor
  Dim pFBuffer As IFeatureBuffer

  Dim pNewFClass As IFeatureClass
  Dim pFields As esriSystem.IVariantArray
  Dim lngSiteIndex As Long
  Dim lngAspectIndex As Long
  Dim lngSlopeIndex As Long
  Dim lngCanopyCoverIndex As Long
  Dim lngBasalAreaIndex As Long
  Dim lngAltBasalAreaIndex As Long
  Dim lngSoilIndex As Long
  Dim lngElevIndex As Long
  Dim lngNorthingIndex As Long
  Dim lngEastingIndex As Long
  Dim lngSpeciesIndex As Long
  Dim lngAbbrevIndex As Long
  Dim lngTypeIndex As Long
  Dim lngLatitudeIndex As Long
  Dim lngLongitudeIndex As Long

  Dim strSpeciesData() As String
  Dim lngSpeciesArrayIndex As Long
  Dim pDoneSpecies As New Collection
  Dim pNewTable As ITable
  Dim pTestWS As IFeatureWorkspace
  Dim pDensityFClass As IFeatureClass
  Dim pCoverFClass As IFeatureClass
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngDensityYearIndex As Long
  Dim lngDensityPlotIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngDensitySpeciesIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverPlotIndex As Long
  Dim lngCoverSiteIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngYearIndex As Long
  Dim lngCommentIndex As Long
  Dim strSpecies As String
  Dim strAbbrev As String
  Dim strType As String
  Dim strSplit() As String
  Dim pRowBuffer As IRowBuffer
  Dim pFeature As IFeature

  Dim pVegComment As Collection
  Set pVegComment = New Collection
  pVegComment.Add "Previously known as Arenaria fendleri; Mat forming perennial forb", "Eremogone fendleri"
  pVegComment.Add "Previously known as Aristida hamulosa", "Aristida ternipes"
  pVegComment.Add "Previously known as Bahia dissecta", "Hymenothrix dissecta"
  pVegComment.Add "Previously known as Blepharoneuron tricholepis", "Muhlenbergia tricholepis"
  pVegComment.Add "Previously known as Lotus wrightii", "Acmispon wrightii"
  pVegComment.Add "Previously known as Bahia dissecta", "Amauriopsis dissecta"
  pVegComment.Add "Previously known as Chamaesyce fendleri", "Euphorbia fendleri"
  pVegComment.Add "Previously known as Chamaesyce revulata", "Euphorbia revoluta"
  pVegComment.Add "Previously known as Chamaesyce serpyllifolia", "Euphorbia serpyllifolia"
  pVegComment.Add "Previously known as Chamaesyce; Could not identify to species level", "Euphorbia sp."
  pVegComment.Add "Previously known as Chenopodium graveolens", "Dysphania graveolens"
  pVegComment.Add "Previously known as Machaeranthera canescens", "Dieteria canescens"
  pVegComment.Add "Previously known as Machaeranthera gracilis", "Xanthisma gracile"
  pVegComment.Add "Previously known as Noccaea montana", "Noccaea fendleri"
  pVegComment.Add "Previously known as Ayenia pusilla", "Ayenia insulicola"
  pVegComment.Add "Mat forming perennial forb", "Antennaria parvifolia"
  pVegComment.Add "Mat forming perennial forb", "Antennaria rosulata"
  pVegComment.Add "Mat forming perennial forb", "Arenaria lanuginosa"
  pVegComment.Add "Could not identify to species level", "Allium sp."
  pVegComment.Add "Could not identify to species level", "Asclepias sp."
  pVegComment.Add "Could not identify to species level", "Astragalus sp."
  pVegComment.Add "Could not identify to species level", "Castilleja sp."
  pVegComment.Add "Could not identify to species level", "Cirsium sp."
  pVegComment.Add "Could not identify to species level", "Erigeron sp."
  pVegComment.Add "Could not identify to species level", "Geranium sp."
  pVegComment.Add "Could not identify to species level", "Linum sp."
  pVegComment.Add "Could not identify to species level", "Lupinus sp."
  pVegComment.Add "Could not identify to species level", "Oxalis sp."
  pVegComment.Add "Could not identify to species level", "Phlox sp."
  pVegComment.Add "Could not identify to species level", "Physaria sp."
  pVegComment.Add "Could not identify to species level", "Potentilla sp."
  pVegComment.Add "Could not identify to species level", "Senecio sp."
  pVegComment.Add "Could not identify to species level", "Solidago sp."
  pVegComment.Add "Could not identify to species level", "Vicia sp."
  pVegComment.Add "Unknown perennial graminoid", "Unknown graminoid"
  Dim strComment As String

  Dim strFinalQuadratList() As String
  Dim pDoneQuadratColl As New Collection
  Dim lngQuadratArrayIndex As Long
  lngQuadratArrayIndex = -1

  lngSpeciesArrayIndex = -1
  If pWS2.NameExists(esriDTTable, "Plant_Species_List") Then
    Set pDataset = pFeatWS.OpenTable("Plant_Species_List")
    pDataset.DELETE
  End If
  If Not pWS2.NameExists(esriDTTable, "Plant_Species_List") Then
    Set pFields = New esriSystem.varArray
    pFields.Add MyGeneralOperations.CreateNewField("Species", esriFieldTypeString, , 255)
    pFields.Add MyGeneralOperations.CreateNewField("Abbreviation", esriFieldTypeString, , 18)
    pFields.Add MyGeneralOperations.CreateNewField("Type", esriFieldTypeString, , 7)
    pFields.Add MyGeneralOperations.CreateNewField("Notes", esriFieldTypeString, , 75)

    Set pNewTable = MyGeneralOperations.CreateGDBTable(pNewWS, "Plant_Species_List", pFields)
    lngSpeciesIndex = pNewTable.FindField("Species")
    lngAbbrevIndex = pNewTable.FindField("Abbreviation")
    lngTypeIndex = pNewTable.FindField("Type")
    lngCommentIndex = pNewTable.FindField("Notes")
    Set pCursor = pNewTable.Insert(True)
    Set pRowBuffer = pNewTable.CreateRowBuffer

    strPurpose = "List of all species observed in all quadrats over all years."

    Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewTable, strAbstract, strPurpose)

    If aml_func_mod.ExistFileDir(strFinalFolder & "\Combined_by_Site.gdb") Then
      Set pTestWS = pNewWSFact.OpenFromFile(strFinalFolder & "\Combined_by_Site.gdb", 0)
    Else
      Set pTestWS = pNewWSFact.OpenFromFile(strFinalFolder & "\Data\Quadrat_Spatial_Data\Combined_by_Site.gdb", 0)
    End If
    Set pDensityFClass = pTestWS.OpenFeatureClass("Density_All")
    lngDensityYearIndex = pDensityFClass.FindField("Year")
    lngDensityPlotIndex = pDensityFClass.FindField("Quadrat")
    lngDensitySiteIndex = pDensityFClass.FindField("Site")
    lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
    Set pCoverFClass = pTestWS.OpenFeatureClass("Cover_All")
    lngCoverYearIndex = pCoverFClass.FindField("Year")
    lngCoverPlotIndex = pCoverFClass.FindField("Quadrat")
    lngCoverSiteIndex = pCoverFClass.FindField("Site")
    lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
    Dim lngCount As Long
    Dim lngCounter As Long

    lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)
    pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
    pProg.position = 0

    lngCounter = 0

    strType = "Density"
    Set pFCursor = pDensityFClass.Search(Nothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      pProg.Step
      lngCounter = lngCounter + 1
      If lngCounter Mod 100 = 0 Then
        DoEvents
      End If
      strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
      If MyGeneralOperations.CheckCollectionForKey(pVegComment, strSpecies) Then
        strComment = pVegComment.Item(strSpecies)
      Else
        strComment = ""
      End If
      strPlot = Trim(pFeature.Value(lngDensityPlotIndex))
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadratColl, strPlot) Then
        lngQuadratArrayIndex = lngQuadratArrayIndex + 1
        ReDim Preserve strFinalQuadratList(lngQuadratArrayIndex)
        strFinalQuadratList(lngQuadratArrayIndex) = strPlot
        pDoneQuadratColl.Add True, strPlot
      End If
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then ' And _
            InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        pDoneSpecies.Add True, strSpecies
        lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
        ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
        strSplit = Split(strSpecies, " ")
        If InStr(1, strSpecies, "No Cover", vbTextCompare) > 0 Then
          strAbbrev = "No Cover Species"
        ElseIf InStr(1, strSpecies, "No Density", vbTextCompare) > 0 Then
          strAbbrev = "No Density Species"
        ElseIf InStr(1, strSpecies, "UNKFOR", vbTextCompare) > 0 Then
          strAbbrev = "Unknown forb"
        ElseIf StrComp(Trim(strSplit(1)), "Sp.", vbTextCompare) = 0 Then
          strAbbrev = UCase(Left(strSplit(0), 4)) & " SP"
        Else
          strAbbrev = UCase(Left(strSplit(0), 3) & Left(strSplit(1), 3))
        End If
        strSpeciesData(0, lngSpeciesArrayIndex) = strSpecies
        strSpeciesData(1, lngSpeciesArrayIndex) = strAbbrev
        strSpeciesData(2, lngSpeciesArrayIndex) = strType
        strSpeciesData(3, lngSpeciesArrayIndex) = strComment
      End If

      Set pFeature = pFCursor.NextFeature
    Loop
    strType = "Cover"
    Set pFCursor = pCoverFClass.Search(Nothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      pProg.Step
      lngCounter = lngCounter + 1
      If lngCounter Mod 100 = 0 Then
        DoEvents
      End If
      strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))
      If MyGeneralOperations.CheckCollectionForKey(pVegComment, strSpecies) Then
        strComment = pVegComment.Item(strSpecies)
      Else
        strComment = ""
      End If
      strPlot = Trim(pFeature.Value(lngCoverPlotIndex))
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadratColl, strPlot) Then
        lngQuadratArrayIndex = lngQuadratArrayIndex + 1
        ReDim Preserve strFinalQuadratList(lngQuadratArrayIndex)
        strFinalQuadratList(lngQuadratArrayIndex) = strPlot
        pDoneQuadratColl.Add True, strPlot
      End If
      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then ' And _
            InStr(1, strSpecies, "Species Observed", vbTextCompare) = 0 Then
        pDoneSpecies.Add True, strSpecies
        lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
        ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
        strSplit = Split(strSpecies, " ")
        If InStr(1, strSpecies, "No Cover", vbTextCompare) > 0 Then
          strAbbrev = "No Cover Species"
        ElseIf InStr(1, strSpecies, "No Density", vbTextCompare) > 0 Then
          strAbbrev = "No Density Species"
        ElseIf StrComp(Trim(strSplit(1)), "Sp.", vbTextCompare) = 0 Then
          strAbbrev = UCase(Left(strSplit(0), 4)) & " SP"
        Else
          strAbbrev = UCase(Left(strSplit(0), 3) & Left(strSplit(1), 3))
        End If
        strSpeciesData(0, lngSpeciesArrayIndex) = strSpecies
        strSpeciesData(1, lngSpeciesArrayIndex) = strAbbrev
        strSpeciesData(2, lngSpeciesArrayIndex) = strType
        strSpeciesData(3, lngSpeciesArrayIndex) = strComment
      End If

      Set pFeature = pFCursor.NextFeature
    Loop

    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, "Dasylirion wheeleri") Then
      pDoneSpecies.Add True, "Dasylirion wheeleri"
      lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
      ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
      strSpeciesData(0, lngSpeciesArrayIndex) = "Dasylirion wheeleri"
      strSpeciesData(1, lngSpeciesArrayIndex) = "DASWHE"
      strSpeciesData(2, lngSpeciesArrayIndex) = "Density"
      strSpeciesData(3, lngSpeciesArrayIndex) = "Observed only in larger 5x5m shrub plots"
    End If
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, "Pinus monophylla") Then
      pDoneSpecies.Add True, "Pinus monophylla"
      lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
      ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
      strSpeciesData(0, lngSpeciesArrayIndex) = "Pinus monophylla"
      strSpeciesData(1, lngSpeciesArrayIndex) = "PINMON"
      strSpeciesData(2, lngSpeciesArrayIndex) = "Density"
      strSpeciesData(3, lngSpeciesArrayIndex) = "Observed only in larger 5x5m shrub plots"
    End If
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, "Rhamnus crocea") Then
      pDoneSpecies.Add True, "Rhamnus crocea"
      lngSpeciesArrayIndex = lngSpeciesArrayIndex + 1
      ReDim Preserve strSpeciesData(3, lngSpeciesArrayIndex)
      strSpeciesData(0, lngSpeciesArrayIndex) = "Rhamnus crocea"
      strSpeciesData(1, lngSpeciesArrayIndex) = "RHACRO"
      strSpeciesData(2, lngSpeciesArrayIndex) = "Density"
      strSpeciesData(3, lngSpeciesArrayIndex) = "Observed only in larger 5x5m shrub plots"
    End If

    QuickSort.StringAscending_TwoDimensional strSpeciesData, 0, lngSpeciesArrayIndex, 0, 3
    For lngIndex = 0 To lngSpeciesArrayIndex
      pRowBuffer.Value(lngSpeciesIndex) = strSpeciesData(0, lngIndex)
      pRowBuffer.Value(lngAbbrevIndex) = strSpeciesData(1, lngIndex)
      pRowBuffer.Value(lngTypeIndex) = strSpeciesData(2, lngIndex)
      pRowBuffer.Value(lngCommentIndex) = strSpeciesData(3, lngIndex)
      pCursor.InsertRow pRowBuffer
    Next lngIndex

    pRowBuffer.Value(lngSpeciesIndex) = "Footnote on Type: 'Cover' = polygon feature for perennial graminoids " & _
        "or mat forming forb, while 'Density' = point feature for annual and perennial forbs, annual graminoids or tree seedlings"
    pRowBuffer.Value(lngAbbrevIndex) = ""
    pRowBuffer.Value(lngTypeIndex) = ""
    pRowBuffer.Value(lngCommentIndex) = ""
    pCursor.InsertRow pRowBuffer

    pCursor.Flush

    ProduceTabularAreaPerSpeciesTable pCoverFClass, lngCoverYearIndex, lngCoverPlotIndex, lngCoverSiteIndex, _
        lngCoverSpeciesIndex, pApp, pSBar, pProg, pCoverFClass.FeatureCount(Nothing), strNewAncillaryFolder, pNewWS, _
        strAbstract, pMxDoc

  Else
    Set pNewTable = pFeatWS.OpenTable("Plant_Species_List")
  End If
  MyGeneralOperations.ExportToCSV pNewTable, strNewAncillaryFolder & "\Plant_Species_List.csv", _
        True, False, False, True, , , True

  Dim pNAD83 As ISpatialReference
  Set pNAD83 = MyGeneralOperations.CreateSpatialReferenceNAD83

  MyGeneralOperations.ExportToCSV_SpecialCases_SA pCoverFClass, strNewAncillaryFolder & "\Cover_Species_Tabular_Version.csv", _
        True, False, False, True, Array("species", "Site", "Quadrat", "Year"), pApp, True, True, pQuadratColl 'pPlotLocColl
  MyGeneralOperations.ExportToCSV_SpecialCases_SA pDensityFClass, strNewAncillaryFolder & "\Density_Species_Tabular_Version.csv", _
        True, False, False, True, Array("species", "Site", "Quadrat", "Year"), pApp, True, True, pQuadratColl 'pPlotLocColl

  FileCopy strSetFolder & "\Summarize_Quadrats_by_Year.csv", strNewAncillaryFolder & "\Summarize_Quadrats_by_Year.csv"
  FileCopy strSetFolder & "\Summarize_by_Site.csv", strNewAncillaryFolder & "\Summarize_by_Site.csv"
  FileCopy strSetFolder & "\Summarize_by_Quadrat.csv", strNewAncillaryFolder & "\Summarize_by_Quadrat.csv"

  Dim booWorked As Boolean
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\Cover_Species_Tabular_Version.csv", pFeatWS)
  Debug.Print "Copy 'Cover_Species_Tabular_Version' Worked = " & CStr(booWorked)
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\Density_Species_Tabular_Version.csv", pFeatWS)
  Debug.Print "Copy 'Density_Species_Tabular_Version' Worked = " & CStr(booWorked)
  booWorked = ExportCSV_ForceFirstRow(strNewAncillaryFolder & "\Summarize_Quadrats_by_Year.csv", pFeatWS)
  Debug.Print "Copy 'Summarize_Quadrats_by_Year' Worked = " & CStr(booWorked)
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\Summarize_by_Site.csv", pFeatWS)
  Debug.Print "Copy 'Summarize_by_Site' Worked = " & CStr(booWorked)
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\Summarize_by_Quadrat.csv", pFeatWS)
  Debug.Print "Copy 'Summarize_by_Quadrat' Worked = " & CStr(booWorked)
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\SA_5x5m_Shrub_Data_and_Quadrat_Locations.csv", pFeatWS)
  Debug.Print "Copy 'SA_5x5m_Shrub_Data_and_Quadrat_Locations' Worked = " & CStr(booWorked)
  booWorked = MyGeneralOperations.CopyCSVTableToGDB(strNewAncillaryFolder & "\Quadrat_Location_Info_Table.csv", pFeatWS)
  Debug.Print "Copy 'Quadrat_Location_Info_Table' Worked = " & CStr(booWorked)

  Debug.Print "Done..."

  pSBar.HideProgressBar
  pProg.position = 0
ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pFinalTable = Nothing
  Set pAddTable = Nothing
  Erase strPlotLocNames
  Set pPlotLocColl = Nothing
  Erase strPlotDataNames
  Set pPlotDataColl = Nothing
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Erase varSites
  Erase varSiteSpecifics
  Erase varArray
  Set pPlotToQuadratColl = Nothing
  Set pQuadratToPlotColl = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Erase strNameSplit
  Set pWStoUpdate = Nothing
  Set pEnumDataset = Nothing
  Set pUpdateDataset = Nothing
  Set pFClass = Nothing
  Set pWS2 = Nothing
  Set pNewFClass = Nothing
  Set pFields = Nothing
  Set pFCursor = Nothing
  Set pFBuffer = Nothing

End Sub

Public Function ExportCSV_ForceFirstRow(strFilename As String, pDestWS As IFeatureWorkspace) As Boolean
  Dim strNewName As String
  Dim strText As String
  Dim strLines() As String
  Dim strLine As String
  Dim strLineSplit() As String

  strText = MyGeneralOperations.ReadTextFile(strFilename)
  strLines = Split(strText, vbCrLf)
  strLine = strLines(0)
  strLineSplit = Split(strLine, ",")
  Dim lngIndex As Long
  Dim strWord As String
  Dim strNewLine As String
  Dim booInQuotes As Boolean

  For lngIndex = 0 To UBound(strLineSplit)
    strWord = strLineSplit(lngIndex)
    booInQuotes = Left(strWord, 1) = """" And Right(strWord, 1) = """"
    strWord = Replace(strWord, """", "")
    strWord = MyGeneralOperations.ReturnAcceptableFieldName2(strWord, Nothing, False, False, False, True)
    Do Until InStr(1, strWord, "__", vbTextCompare) = 0
      strWord = Replace(strWord, "__", "_")
    Loop
    If Right(strWord, 1) = "_" Then strWord = Left(strWord, Len(strWord) - 1)
    If Left(strWord, 1) = "_" Then strWord = Right(strWord, Len(strWord) - 1)
    If booInQuotes Then strWord = """" & strWord & """"
    strNewLine = strNewLine & strWord & IIf(lngIndex = UBound(strLineSplit), "", ",")
  Next lngIndex

  Dim strNewCSV As String
  strNewCSV = strNewLine & vbCrLf
  For lngIndex = 1 To UBound(strLines)
    strNewCSV = strNewCSV & strLines(lngIndex) & IIf(lngIndex = UBound(strLines), "", vbCrLf)
  Next lngIndex

  Dim strTempPath As String
  strTempPath = aml_func_mod.ReturnDir3(strFilename, True) & "temp.csv"
  MyGeneralOperations.WriteTextFile strTempPath, strNewCSV, True, False

  ExportCSV_ForceFirstRow = MyGeneralOperations.CopyCSVTableToGDB(strTempPath, pDestWS, "Summarize_Quadrats_by_Year")
  Kill strTempPath

End Function

Public Sub ProduceTabularAreaPerSpeciesTable(pCoverFClass As IFeatureClass, lngCoverYearIndex As Long, _
    lngCoverPlotIndex As Long, lngCoverSiteIndex As Long, lngCoverSpeciesIndex As Long, pApp As IApplication, _
    pSBar As IStatusBar, pProg As IStepProgressor, lngCount As Long, strNewAncillaryFolder As String, _
    pNewWS As IWorkspace, strAbstract As String, pMxDoc As IMxDocument)

  Dim lngCounter As Long

  pSBar.ShowProgressBar "Cover Basal Area Pass 1 of 2...", 0, lngCount, 1, True
  pProg.position = 0

  lngCounter = 0
  Dim strSortArray() As String
  Dim lngSortCounter As Long
  Dim pDoneColl As New Collection
  lngSortCounter = -1
  Dim strSpecies As String
  Dim strSite As String
  Dim strYear As String
  Dim strQuadrat As String

  Dim strPrefix As String
  Dim strSuffix As String
  Dim strKey As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pCoverFClass, strPrefix, strSuffix)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature

  Dim strType As String
  strType = "Cover"
  Dim strBaseQuery As String

  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSite = Trim(pFeature.Value(lngCoverSiteIndex))
    strQuadrat = Trim(pFeature.Value(lngCoverPlotIndex))
    strKey = strSite & ":" & strQuadrat
    If Not MyGeneralOperations.CheckCollectionForKey(pDoneColl, strKey) Then
      pDoneColl.Add True, strKey
      lngSortCounter = lngSortCounter + 1
      ReDim Preserve strSortArray(2, lngSortCounter)
      strSortArray(0, lngSortCounter) = strSite
      strSortArray(1, lngSortCounter) = strQuadrat
      strSortArray(2, lngSortCounter) = strPrefix & "Site" & strSuffix & " = '" & strSite & "' AND " & _
          strPrefix & "Quadrat" & strSuffix & " = '" & strQuadrat & "'"
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  QuickSort.StringAscending_TwoDimensional strSortArray, 0, lngSortCounter, 0, 2
  Dim lngIndex As Long

  Dim strYears() As String
  Dim strSpeciesArray() As String
  Dim lngYearIndex As Long
  Dim lngSpeciesIndex As Long
  Dim strBaseYearQuery As String
  Dim strBaseSpeciesQuery As String
  Dim dblTotalArea As Double
  Dim lngObservationCount As Long

  Dim pFields As esriSystem.IVariantArray
  Set pFields = New esriSystem.varArray
  pFields.Add MyGeneralOperations.CreateNewField("Site", esriFieldTypeString, , 35)
  pFields.Add MyGeneralOperations.CreateNewField("Quadrat", esriFieldTypeString, , 15)
  pFields.Add MyGeneralOperations.CreateNewField("Year", esriFieldTypeString, , 5)
  pFields.Add MyGeneralOperations.CreateNewField("Type", esriFieldTypeString, , 5)
  pFields.Add MyGeneralOperations.CreateNewField("Species", esriFieldTypeString, , 50)
  pFields.Add MyGeneralOperations.CreateNewField("Number_Observations", esriFieldTypeInteger)
  pFields.Add MyGeneralOperations.CreateNewField("Area_Sq_Cm", esriFieldTypeDouble)
  pFields.Add MyGeneralOperations.CreateNewField("Proportion_Quadrat", esriFieldTypeString, , 15)

  Dim pNewTable As ITable

  Dim pWS2 As IWorkspace2
  Dim pFeatWS As IFeatureWorkspace
  Set pWS2 = pNewWS
  Set pFeatWS = pNewWS
  Dim pDataset As IDataset
  If pWS2.NameExists(esriDTTable, "Basal_Cover_per_Species_by_Quadrat_and_Year") Then
    Set pDataset = pFeatWS.OpenTable("Basal_Cover_per_Species_by_Quadrat_and_Year")
    pDataset.DELETE
  End If

  Set pNewTable = MyGeneralOperations.CreateGDBTable(pNewWS, "Basal_Cover_per_Species_by_Quadrat_and_Year", pFields)

  Dim lngNewSiteIndex As Long
  Dim lngNewQuadratIndex As Long
  Dim lngNewYearIndex As Long
  Dim lngNewTypeIndex As Long
  Dim lngNewSpeciesIndex As Long
  Dim lngNewObsCountIndex As Long
  Dim lngNewAreaIndex As Long
  Dim lngNewProportionIndex As Long

  lngNewSiteIndex = pNewTable.FindField("Site")
  lngNewQuadratIndex = pNewTable.FindField("Quadrat")
  lngNewYearIndex = pNewTable.FindField("Year")
  lngNewTypeIndex = pNewTable.FindField("Type")
  lngNewSpeciesIndex = pNewTable.FindField("Species")
  lngNewObsCountIndex = pNewTable.FindField("Number_Observations")
  lngNewAreaIndex = pNewTable.FindField("Area_Sq_Cm")
  lngNewProportionIndex = pNewTable.FindField("Proportion_Quadrat")

  Dim strPurpose As String
  strPurpose = "List of all species observed in all quadrats over all years."

  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewTable, strAbstract, strPurpose)

  Dim pRowBuffer As IRowBuffer
  Dim pCursor As ICursor
  Set pCursor = pNewTable.Insert(True)
  Set pRowBuffer = pNewTable.CreateRowBuffer

  pSBar.ShowProgressBar "Cover Basal Area Pass 2 of 2...", 0, lngSortCounter, 1, True
  pProg.position = 0

  For lngIndex = 0 To lngSortCounter
    pProg.Step
    DoEvents
    strSite = strSortArray(0, lngIndex)
    strQuadrat = strSortArray(1, lngIndex)
    strBaseQuery = strSortArray(2, lngIndex)

    strYears = ReturnArrayOfValues(pCoverFClass, lngCoverYearIndex, strBaseQuery)
    For lngYearIndex = 0 To UBound(strYears, 2)
      strYear = strYears(0, lngYearIndex)
      strBaseYearQuery = strBaseQuery & " AND " & strPrefix & "Year" & strSuffix & " = '" & strYear & "'"

      strSpeciesArray = ReturnArrayOfValues(pCoverFClass, lngCoverSpeciesIndex, strBaseYearQuery)
      For lngSpeciesIndex = 0 To UBound(strSpeciesArray, 2)
        strSpecies = strSpeciesArray(0, lngSpeciesIndex)
        strBaseSpeciesQuery = strBaseYearQuery & " AND " & strPrefix & "Species" & strSuffix & " = '" & strSpecies & "'"

        FillCountAndAreaForSpecies pCoverFClass, strBaseSpeciesQuery, lngObservationCount, dblTotalArea

        pRowBuffer.Value(lngNewSiteIndex) = strSite
        pRowBuffer.Value(lngNewQuadratIndex) = strQuadrat
        pRowBuffer.Value(lngNewYearIndex) = strYear
        pRowBuffer.Value(lngNewTypeIndex) = strType
        pRowBuffer.Value(lngNewSpeciesIndex) = strSpecies
        pRowBuffer.Value(lngNewObsCountIndex) = lngObservationCount
        pRowBuffer.Value(lngNewAreaIndex) = dblTotalArea
        pRowBuffer.Value(lngNewProportionIndex) = Format(dblTotalArea / 10000, "0.00%")
        pCursor.InsertRow pRowBuffer

      Next lngSpeciesIndex
    Next lngYearIndex
    pCursor.Flush
  Next lngIndex

  pCursor.Flush

  MyGeneralOperations.ExportToCSV pNewTable, strNewAncillaryFolder & "\Basal_Cover_per_Species_by_Quadrat_and_Year.csv", _
        True, False, False, True, , pApp, True

ClearMemory:
  Erase strSortArray
  Set pDoneColl = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strYears
  Erase strSpeciesArray
  Set pFields = Nothing
  Set pNewTable = Nothing
  Set pRowBuffer = Nothing
  Set pCursor = Nothing

End Sub

Public Sub FillCountAndAreaForSpecies(pCoverFClass As IFeatureClass, strQueryString As String, lngObservationCount As Long, _
    dblTotalSqCm As Double)

  Dim pQueryFilt As IQueryFilter
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pPoly As IPolygon
  Dim pArea As IArea
  Dim dblCumulativeArea As Double
  Dim lngCounter As Long

  lngCounter = 0
  dblCumulativeArea = 0

  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString
  Set pFCursor = pCoverFClass.Search(pQueryFilt, True)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pPoly = pFeature.Shape
    Set pArea = pPoly
    dblCumulativeArea = dblCumulativeArea + (pArea.Area * 10000)
    lngCounter = lngCounter + 1
    Set pFeature = pFCursor.NextFeature
  Loop

  dblTotalSqCm = dblCumulativeArea
  lngObservationCount = lngCounter

ClearMemory:
  Set pQueryFilt = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pPoly = Nothing
  Set pArea = Nothing

End Sub

Public Function ReturnArrayOfValues(pCoverFClass As IFeatureClass, lngFieldIndex As Long, strQueryString As String) As String()

  Dim pQueryFilt As IQueryFilter
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strReturn() As String
  Dim lngArrayIndex As Long
  Dim pDoneColl As New Collection
  Dim lngCounter As Long
  Dim strValue As String

  lngArrayIndex = -1
  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString
  Set pFCursor = pCoverFClass.Search(pQueryFilt, True)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strValue = pFeature.Value(lngFieldIndex)
    If MyGeneralOperations.CheckCollectionForKey(pDoneColl, strValue) Then
      lngCounter = pDoneColl.Item(strValue)
      pDoneColl.Remove strValue
    Else
      lngCounter = 0
      lngArrayIndex = lngArrayIndex + 1
      ReDim Preserve strReturn(1, lngArrayIndex)
      strReturn(0, lngArrayIndex) = strValue
    End If
    lngCounter = lngCounter + 1
    pDoneColl.Add lngCounter, strValue
    Set pFeature = pFCursor.NextFeature
  Loop

  Dim lngIndex As Long
  For lngIndex = 0 To lngArrayIndex
    strValue = strReturn(0, lngArrayIndex)
    strReturn(1, lngArrayIndex) = pDoneColl.Item(strValue)
  Next lngIndex

  QuickSort.StringAscending_TwoDimensional strReturn, 0, lngArrayIndex, 0, 1

  ReturnArrayOfValues = strReturn

ClearMemory:
  Set pQueryFilt = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strReturn
  Set pDoneColl = Nothing

End Function

Public Sub DeclareWorkspaces(strOrigShapefiles As String, Optional strModifiedRoot As String, _
    Optional strShiftedRoot As String, Optional strExportBase As String, Optional strRecreatedModifiedRoot As String, _
    Optional strSetFolder As String, Optional strExtractShapefileFolder As String, Optional strFinalFolder As String, _
    Optional datDateUsed As Date)

  Dim booUseCurrentDate As Boolean
  booUseCurrentDate = False

  Dim strSpecifiedDate As String
  strSpecifiedDate = "2024_12_22"

  Dim strDate As String
  Dim strDateSplit() As String
  Dim strCurrentDate As String

  If booUseCurrentDate Then
    datDateUsed = Now
    strDate = Replace(Format(Now, "yyyy_mm_dd"), "Sep_", "Sept_")   ' "2021_06_08"
    strCurrentDate = Replace(Format(Now, "mmm_d_yyyy"), "Sep_", "Sept_")
  Else
    strDate = strSpecifiedDate
    strDateSplit = Split(strDate, "_")

    datDateUsed = DateSerial(CInt(strDateSplit(0)), CInt(strDateSplit(1)), CInt(strDateSplit(2)))
    strCurrentDate = Format(DateSerial(CInt(strDateSplit(0)), CInt(strDateSplit(1)), CInt(strDateSplit(2))), "mmm_d_yyyy")
    strCurrentDate = Replace(strCurrentDate, "Sep_", "Sept_")
  End If

  strOrigShapefiles = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\contemporary_data_" & strDate
  strModifiedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Modified_Data_" & strDate
  strRecreatedModifiedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Cleaned_Data_" & strDate
  strShiftedRoot = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Cleaned_Data_" & strDate & "_Shift"
  strExportBase = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Map_Exports_" & strDate
  strSetFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate
  strExtractShapefileFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Species_Shapefile_Extractions"
  strFinalFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\" & strDate & "\Final_Datasets_" & strCurrentDate

End Sub

Public Function ReturnMissingShapefiles(varCheckArray() As Variant) As String
  Dim lngIndex As Long
  Dim strCheck As String
  Dim strReport As String
  For lngIndex = 0 To UBound(varCheckArray, 2)
    If varCheckArray(1, lngIndex) = False Then
      strCheck = varCheckArray(0, lngIndex)
      strReport = strReport & CStr(lngIndex) & "] " & strCheck & vbCrLf
    End If
  Next lngIndex

  ReturnMissingShapefiles = strReport
End Function

Public Sub UpdateCheckArray(varCheckArray() As Variant, strPath As String)
  Dim lngIndex As Long
  Dim strCheck As String
  For lngIndex = 0 To UBound(varCheckArray, 2)
    strCheck = varCheckArray(0, lngIndex)
    If InStr(1, strPath, strCheck, vbTextCompare) > 0 Then
      varCheckArray(1, lngIndex) = True
    End If
  Next lngIndex

End Sub

Public Function BuildCheckArray_SA(strAllPaths() As String) As Variant()

  Dim varReturn() As Variant
  Dim lngCounter As Long

  lngCounter = -1

  Dim lngIndex As Long
  Dim strVal As String
  For lngIndex = 0 To UBound(strAllPaths)
    strVal = strAllPaths(lngIndex)
    If StrComp(Right(strVal, 4), ".dbf", vbTextCompare) = 0 Then
      lngCounter = lngCounter + 1
      ReDim Preserve varReturn(1, lngCounter)
      strVal = aml_func_mod.ReturnFilename2(strVal)
      strVal = aml_func_mod.ClipExtension2(strVal)
      varReturn(0, lngCounter) = strVal
      varReturn(1, lngCounter) = False
    End If
  Next lngIndex

  BuildCheckArray_SA = varReturn

End Function

Public Sub FillNameConverters_SA_Original(varNameLinks() As Variant, p2020toOld As Collection, pOldTo2020 As Collection)

  ReDim varNameLinks(541)

  varNameLinks(0) = Array("Q10_1935_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(1) = Array("Q10_1935_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(2) = Array("Q11_1935_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(3) = Array("Q11_1935_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(4) = Array("Q12_1935_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(5) = Array("Q12_1935_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(6) = Array("Q13_1935_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(7) = Array("Q13_1935_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(8) = Array("Q14_1935_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(9) = Array("Q14_1935_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(10) = Array("Q15_1935_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(11) = Array("Q15_1935_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(12) = Array("Q16_1935_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(13) = Array("Q16_1935_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(14) = Array("Q17_1935_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(15) = Array("Q17_1935_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(16) = Array("Q18_1935_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(17) = Array("Q18_1935_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(18) = Array("Q19_1935_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(19) = Array("Q19_1935_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(20) = Array("Q1_1935_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(21) = Array("Q1_1935_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(22) = Array("Q20_1935_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(23) = Array("Q20_1935_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(24) = Array("Q21_1935_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(25) = Array("Q21_1935_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(26) = Array("Q22_1935_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(27) = Array("Q22_1935_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(28) = Array("Q23_1935_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(29) = Array("Q23_1935_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(30) = Array("Q24_1935_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(31) = Array("Q24_1935_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(32) = Array("Q2_1935_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(33) = Array("Q2_1935_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(34) = Array("Q3_1935_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(35) = Array("Q3_1935_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(36) = Array("Q4_1935_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(37) = Array("Q4_1935_D", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(38) = Array("Q5_1395_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(39) = Array("Q5_1935_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(40) = Array("Q6_1935_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(41) = Array("Q6_1935_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(42) = Array("Q7_1935_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(43) = Array("Q7_1935_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(44) = Array("Q8_1935_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(45) = Array("Q8_1935_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(46) = Array("Q9_1935_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(47) = Array("Q9_1935_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(48) = Array("Q10_1936_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(49) = Array("Q10_1936_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(50) = Array("Q11_1936_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(51) = Array("Q11_1936_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(52) = Array("Q12_1936_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(53) = Array("Q12_1936_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(54) = Array("Q13_1936_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(55) = Array("Q13_1936_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(56) = Array("Q14_1936_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(57) = Array("Q14_1936_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(58) = Array("Q15_1936_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(59) = Array("Q15_1936_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(60) = Array("Q16_1936_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(61) = Array("Q16_1936_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(62) = Array("Q17_1936_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(63) = Array("Q17_1936_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(64) = Array("Q18_1936_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(65) = Array("Q18_1936_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(66) = Array("Q19_1936_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(67) = Array("Q19_1936_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(68) = Array("Q1_1936_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(69) = Array("Q1_1936_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(70) = Array("Q20_1936_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(71) = Array("Q20_1936_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(72) = Array("Q21_1936_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(73) = Array("Q21_1936_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(74) = Array("Q22_1936_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(75) = Array("Q22_1936_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(76) = Array("Q23_1936_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(77) = Array("Q23_1936_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(78) = Array("Q24_1936_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(79) = Array("Q24_1936_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(80) = Array("Q2_1936_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(81) = Array("Q2_1936_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(82) = Array("Q3_1936_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(83) = Array("Q3_1936_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(84) = Array("Q4_1936_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(85) = Array("Q4_1936_D", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(86) = Array("Q5_1936_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(87) = Array("Q5_1936_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(88) = Array("Q6_1936_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(89) = Array("Q6_1936_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(90) = Array("Q7_1936_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(91) = Array("Q7_1936_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(92) = Array("Q8_1936_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(93) = Array("Q8_1936_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(94) = Array("Q9_1936_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(95) = Array("Q9_1936_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(96) = Array("Q10_1940_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(97) = Array("Q10_1940_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(98) = Array("Q11_1940_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(99) = Array("Q11_1940_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(100) = Array("Q12_1940_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(101) = Array("Q12_1940_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(102) = Array("Q13_1940_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(103) = Array("Q13_1940_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(104) = Array("Q14_1940_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(105) = Array("Q14_1940_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(106) = Array("Q15_1940_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(107) = Array("Q15_1940_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(108) = Array("Q16_1940_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(109) = Array("Q16_1940_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(110) = Array("Q17_1940_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(111) = Array("Q17_1940_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(112) = Array("Q18_1940_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(113) = Array("Q18_1940_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(114) = Array("Q19_1940_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(115) = Array("Q19_1940_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(116) = Array("Q1_1940_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(117) = Array("Q1_1940_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(118) = Array("Q20_1940_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(119) = Array("Q20_1940_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(120) = Array("Q23_1940_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(121) = Array("Q23_1940_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(122) = Array("Q24_1940_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(123) = Array("Q24_1940_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(124) = Array("Q2_1940_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(125) = Array("Q2_1940_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(126) = Array("Q3_1940_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(127) = Array("Q3_1940_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(128) = Array("Q4_1940_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(129) = Array("Q4_1940_D", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(130) = Array("Q5_1940_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(131) = Array("Q5_1940_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(132) = Array("Q6_1940_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(133) = Array("Q6_1940_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(134) = Array("Q7_1940_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(135) = Array("Q7_1940_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(136) = Array("Q8_1940_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(137) = Array("Q8_1940_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(138) = Array("Q9_1940_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(139) = Array("Q9_1940_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(140) = Array("Q21_1941_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(141) = Array("Q21_1941_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(142) = Array("Q22_1941_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(143) = Array("Q22_1941_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(144) = Array("Q10_1942_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(145) = Array("Q10_1942_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(146) = Array("Q11_1942_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(147) = Array("Q11_1942_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(148) = Array("Q12_1942_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(149) = Array("Q12_1942_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(150) = Array("Q13_1942_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(151) = Array("Q13_1942_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(152) = Array("Q14_1942_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(153) = Array("Q14_1942_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(154) = Array("Q15_1942_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(155) = Array("Q15_1942_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(156) = Array("Q16_1942_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(157) = Array("Q16_1942_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(158) = Array("Q17_1942_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(159) = Array("Q17_1942_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(160) = Array("Q18_1942_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(161) = Array("Q18_1942_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(162) = Array("Q19_1942_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(163) = Array("Q19_1942_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(164) = Array("Q1_1942_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(165) = Array("Q1_1942_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(166) = Array("Q20_1942_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(167) = Array("Q20_1942_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(168) = Array("Q21_1942_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(169) = Array("Q21_1942_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(170) = Array("Q22_1942_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(171) = Array("Q22_1942_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(172) = Array("Q23_1942_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(173) = Array("Q23_1942_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(174) = Array("Q24_1942_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(175) = Array("Q24_1942_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(176) = Array("Q2_1942_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(177) = Array("Q2_1942_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(178) = Array("Q3_1942_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(179) = Array("Q3_1942_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(180) = Array("Q4_1942_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(181) = Array("Q4_1942_D", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(182) = Array("Q5_1942_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(183) = Array("Q5_1942_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(184) = Array("Q6_1942_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(185) = Array("Q6_1942_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(186) = Array("Q7_1942_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(187) = Array("Q7_1942_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(188) = Array("Q8_1942_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(189) = Array("Q8_1942_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(190) = Array("Q9_1942_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(191) = Array("Q9_1942_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(192) = Array("Q10_1952_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(193) = Array("Q10_1952_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(194) = Array("Q11_1952_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(195) = Array("Q11_1952_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(196) = Array("Q12_1952_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(197) = Array("Q12_1952_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(198) = Array("Q13_1952_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(199) = Array("Q13_1952_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(200) = Array("Q14_1952_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(201) = Array("Q14_1952_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(202) = Array("Q15_1952_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(203) = Array("Q15_1952_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(204) = Array("Q16_1952_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(205) = Array("Q16_1952_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(206) = Array("Q17_1952_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(207) = Array("Q17_1952_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(208) = Array("Q18_1952_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(209) = Array("Q18_1952_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(210) = Array("Q19_1952_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(211) = Array("Q19_1952_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(212) = Array("Q1_1952_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(213) = Array("Q1_1952_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(214) = Array("Q20_1952_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(215) = Array("Q20_1952_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(216) = Array("Q21_1952_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(217) = Array("Q21_1952_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(218) = Array("Q22_1952_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(219) = Array("Q22_1952_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(220) = Array("Q23_1952_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(221) = Array("Q23_1952_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(222) = Array("Q24_1952_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(223) = Array("Q24_1952_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(224) = Array("Q2_1952_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(225) = Array("Q2_1952_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(226) = Array("Q3_1952_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(227) = Array("Q3_1952_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(228) = Array("Q4_1952_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(229) = Array("Q4_1952_D", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(230) = Array("Q5_1952_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(231) = Array("Q5_1952_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(232) = Array("Q6_1952_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(233) = Array("Q6_1952_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(234) = Array("Q7_1952_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(235) = Array("Q7_1952_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(236) = Array("Q8_1952_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(237) = Array("Q8_1952_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(238) = Array("Q9_1952_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(239) = Array("Q9_1952_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(240) = Array("Q10_1955_C", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(241) = Array("Q10_1955_D", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(242) = Array("Q11_1955_C", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(243) = Array("Q11_1955_D", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(244) = Array("Q12_1955_C", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(245) = Array("Q12_1955_D", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(246) = Array("Q13_1955_C", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(247) = Array("Q13_1955_D", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(248) = Array("Q13_1955_DCopy", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(249) = Array("Q14_1955_C", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(250) = Array("Q14_1955_D", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(251) = Array("Q15_1955_C", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(252) = Array("Q15_1955_D", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(253) = Array("Q16_1955_C", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(254) = Array("Q16_1955_D", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(255) = Array("Q17_1955_C", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(256) = Array("Q17_1955_D", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(257) = Array("Q18_1955_C", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(258) = Array("Q18_1955_D", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(259) = Array("Q19_1955_C", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(260) = Array("Q19_1955_D", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(261) = Array("Q1_1955_C", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(262) = Array("Q1_1955_D", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(263) = Array("Q20_1955_C", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(264) = Array("Q20_1955_D", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(265) = Array("Q21_1955_C", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(266) = Array("Q21_1955_D", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(267) = Array("Q22_1955_C", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(268) = Array("Q22_1955_D", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(269) = Array("Q23_1955_C", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(270) = Array("Q23_1955_D", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(271) = Array("Q24_1955_C", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(272) = Array("Q24_1955_D", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(273) = Array("Q2_1955_C", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(274) = Array("Q2_1955_D", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(275) = Array("Q3_1955_C", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(276) = Array("Q3_1955_D", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(277) = Array("Q4_1955_C", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(278) = Array("Q5_1955_C", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(279) = Array("Q5_1955_D", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(280) = Array("Q6_1955_C", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(281) = Array("Q6_1955_D", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(282) = Array("Q7_1955_C", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(283) = Array("Q7_1955_D", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(284) = Array("Q8_1955_C", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(285) = Array("Q8_1955_D", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(286) = Array("Q9_1955_C", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(287) = Array("Q9_1955_D", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(288) = Array("Q_ND_10_2017_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(289) = Array("Q_ND_10_2017_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(290) = Array("Q_ND_11_2017_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(291) = Array("Q_ND_11_2017_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(292) = Array("Q_ND_12_2017_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(293) = Array("Q_ND_12_2017_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(294) = Array("Q_ND_13_2017_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(295) = Array("Q_ND_13_2017_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(296) = Array("Q_ND_14_2017_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(297) = Array("Q_ND_14_2017_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(298) = Array("Q_ND_15_2017_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(299) = Array("Q_ND_15_2017_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(300) = Array("Q_ND_16_2017_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(301) = Array("Q_ND_16_2017_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(302) = Array("Q_ND_17_2017_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(303) = Array("Q_ND_17_2017_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(304) = Array("Q_ND_18_2017_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(305) = Array("Q_ND_18_2017_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(306) = Array("Q_ND_19_2017_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(307) = Array("Q_ND_19_2017_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(308) = Array("Q_ND_1_2017_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(309) = Array("Q_ND_1_2017_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(310) = Array("Q_ND_20_2017_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(311) = Array("Q_ND_20_2017_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(312) = Array("Q_ND_21_2017_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(313) = Array("Q_ND_21_2017_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(314) = Array("Q_ND_22_2017_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(315) = Array("Q_ND_22_2017_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(316) = Array("Q_ND_23_2017_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(317) = Array("Q_ND_23_2017_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(318) = Array("Q_ND_24_2017_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(319) = Array("Q_ND_24_2017_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(320) = Array("Q_ND_2_2017_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(321) = Array("Q_ND_2_2017_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(322) = Array("Q_ND_3_2017_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(323) = Array("Q_ND_3_2017_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(324) = Array("Q_ND_4_2017_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(325) = Array("Q_ND_4_2017_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(326) = Array("Q_ND_5_2017_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(327) = Array("Q_ND_5_2017_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(328) = Array("Q_ND_6_2017_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(329) = Array("Q_ND_6_2017_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(330) = Array("Q_ND_7_2017_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(331) = Array("Q_ND_7_2017_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(332) = Array("Q_ND_8_2017_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(333) = Array("Q_ND_8_2017_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(334) = Array("Q_ND_9_2017_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(335) = Array("Q_ND_9_2017_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(336) = Array("Q11974_2017_C", "Plot_3_Q11974")
  varNameLinks(337) = Array("Q11974_2017_D", "Plot_3_Q11974")
  varNameLinks(338) = Array("Q11975_2017_C", "Plot_3_Q11975")
  varNameLinks(339) = Array("Q11975_2017_D", "Plot_3_Q11975")
  varNameLinks(340) = Array("Q11977_2017_C", "Plot_3_Q11977")
  varNameLinks(341) = Array("Q11977_2017_D", "Plot_3_Q11977")
  varNameLinks(342) = Array("Q33401_2017_C", "Plot_3_Q33401")
  varNameLinks(343) = Array("Q33401_2017_D", "Plot_3_Q33401")
  varNameLinks(344) = Array("Q33402_2017_C", "Plot_3_Q33402")
  varNameLinks(345) = Array("Q33402_2017_D", "Plot_3_Q33402")
  varNameLinks(346) = Array("Q33404_2017_C", "Plot_3_Q33404")
  varNameLinks(347) = Array("Q33404_2017_D", "Plot_3_Q33404")
  varNameLinks(348) = Array("Q_ND_10_2018_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(349) = Array("Q_ND_10_2018_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(350) = Array("Q_ND_11_2018_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(351) = Array("Q_ND_11_2018_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(352) = Array("Q_ND_12_2018_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(353) = Array("Q_ND_12_2018_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(354) = Array("Q_ND_13_2018_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(355) = Array("Q_ND_13_2018_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(356) = Array("Q_ND_14_2018_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(357) = Array("Q_ND_14_2018_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(358) = Array("Q_ND_15_2018_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(359) = Array("Q_ND_15_2018_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(360) = Array("Q_ND_16_2018_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(361) = Array("Q_ND_16_2018_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(362) = Array("Q_ND_17_2018_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(363) = Array("Q_ND_17_2018_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(364) = Array("Q_ND_18_2018_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(365) = Array("Q_ND_18_2018_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(366) = Array("Q_ND_19_2018_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(367) = Array("Q_ND_19_2018_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(368) = Array("Q_ND_1_2018_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(369) = Array("Q_ND_1_2018_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(370) = Array("Q_ND_20_2018_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(371) = Array("Q_ND_20_2018_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(372) = Array("Q_ND_21_2018_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(373) = Array("Q_ND_21_2018_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(374) = Array("Q_ND_22_2018_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(375) = Array("Q_ND_22_2018_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(376) = Array("Q_ND_23_2018_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(377) = Array("Q_ND_23_2018_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(378) = Array("Q_ND_24_2018_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(379) = Array("Q_ND_24_2018_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(380) = Array("Q_ND_2_2018_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(381) = Array("Q_ND_2_2018_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(382) = Array("Q_ND_3_2018_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(383) = Array("Q_ND_3_2018_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(384) = Array("Q_ND_4_2018_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(385) = Array("Q_ND_4_2018_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(386) = Array("Q_ND_5_2018_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(387) = Array("Q_ND_5_2018_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(388) = Array("Q_ND_6_2018_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(389) = Array("Q_ND_6_2018_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(390) = Array("Q_ND_7_2018_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(391) = Array("Q_ND_7_2018_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(392) = Array("Q_ND_8_2018_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(393) = Array("Q_ND_8_2018_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(394) = Array("Q_ND_9_2018_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(395) = Array("Q_ND_9_2018_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(396) = Array("QP3_11974_2018_CF", "Plot_3_Q11974")
  varNameLinks(397) = Array("QP3_11974_2018_DF", "Plot_3_Q11974")
  varNameLinks(398) = Array("QP3_11975_2018_CF", "Plot_3_Q11975")
  varNameLinks(399) = Array("QP3_11975_2018_DF", "Plot_3_Q11975")
  varNameLinks(400) = Array("QP3_11977_2018_CF", "Plot_3_Q11977")
  varNameLinks(401) = Array("QP3_11977_2018_DF", "Plot_3_Q11977")
  varNameLinks(402) = Array("QP3_33401_2018_C", "Plot_3_Q33401")
  varNameLinks(403) = Array("QP3_33401_2018_D", "Plot_3_Q33401")
  varNameLinks(404) = Array("QP3_33402_2018_C", "Plot_3_Q33402")
  varNameLinks(405) = Array("QP3_33402_2018_D", "Plot_3_Q33402")
  varNameLinks(406) = Array("QP3_33403_2018_C", "Plot_3_Q33403")
  varNameLinks(407) = Array("QP3_33403_2018_D", "Plot_3_Q33403")
  varNameLinks(408) = Array("Q_ND_10_2019_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(409) = Array("Q_ND_10_2019_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(410) = Array("Q_ND_11_2019_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(411) = Array("Q_ND_11_2019_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(412) = Array("Q_ND_12_2019_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(413) = Array("Q_ND_12_2019_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(414) = Array("Q_ND_13_2019_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(415) = Array("Q_ND_13_2019_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(416) = Array("Q_ND_14_2019_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(417) = Array("Q_ND_14_2019_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(418) = Array("Q_ND_15_2019_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(419) = Array("Q_ND_15_2019_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(420) = Array("Q_ND_16_2019_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(421) = Array("Q_ND_16_2019_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(422) = Array("Q_ND_17_2019_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(423) = Array("Q_ND_17_2019_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(424) = Array("Q_ND_18_2019_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(425) = Array("Q_ND_18_2019_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(426) = Array("Q_ND_19_2019_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(427) = Array("Q_ND_19_2019_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(428) = Array("Q_ND_1_2019_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(429) = Array("Q_ND_1_2019_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(430) = Array("Q_ND_20_2019_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(431) = Array("Q_ND_20_2019_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(432) = Array("Q_ND_21_2019_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(433) = Array("Q_ND_21_2019_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(434) = Array("Q_ND_22_2019_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(435) = Array("Q_ND_22_2019_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(436) = Array("Q_ND_23_2019_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(437) = Array("Q_ND_23_2019_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(438) = Array("Q_ND_24_2019_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(439) = Array("Q_ND_24_2019_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(440) = Array("Q_ND_2_2019_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(441) = Array("Q_ND_2_2019_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(442) = Array("Q_ND_3_2019_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(443) = Array("Q_ND_3_2019_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(444) = Array("Q_ND_4_2019_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(445) = Array("Q_ND_4_2019_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(446) = Array("Q_ND_5_2019_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(447) = Array("Q_ND_5_2019_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(448) = Array("Q_ND_6_2019_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(449) = Array("Q_ND_6_2019_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(450) = Array("Q_ND_7_2019_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(451) = Array("Q_ND_7_2019_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(452) = Array("Q_ND_8_2019_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(453) = Array("Q_ND_8_2019_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(454) = Array("Q_ND_9_2019_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(455) = Array("Q_ND_9_2019_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(456) = Array("QP1_11969_2019_CF", "Plot_1_Q11969")
  varNameLinks(457) = Array("QP1_11969_2019_DF", "Plot_1_Q11969")
  varNameLinks(458) = Array("QP1_387_2019_CF", "Plot_1_Q387")
  varNameLinks(459) = Array("QP1_387_2019_DF", "Plot_1_Q387")
  varNameLinks(460) = Array("QP1_390_2019_CF", "Plot_1_Q390")
  varNameLinks(461) = Array("QP1_390_2019_DF", "Plot_1_Q390")
  varNameLinks(462) = Array("Q_P3_11974_2019_CF", "Plot_3_Q11974")
  varNameLinks(463) = Array("Q_P3_11974_2019_DF", "Plot_3_Q11974")
  varNameLinks(464) = Array("Q_P3_11975_2019_CF", "Plot_3_Q11975")
  varNameLinks(465) = Array("Q_P3_11975_2019_DF", "Plot_3_Q11975")
  varNameLinks(466) = Array("Q_P3_11977_2019_CF", "Plot_3_Q11977")
  varNameLinks(467) = Array("Q_P3_11977_2019_DF", "Plot_3_Q11977")
  varNameLinks(468) = Array("Q_P3_33401_2019_CF", "Plot_3_Q33401")
  varNameLinks(469) = Array("Q_P3_33401_2019_DF", "Plot_3_Q33401")
  varNameLinks(470) = Array("Q_P3_33402_2019_CF", "Plot_3_Q33402")
  varNameLinks(471) = Array("Q_P3_33402_2019_DF", "Plot_3_Q33402")
  varNameLinks(472) = Array("Q_P3_33403_2019_CF", "Plot_3_Q33403")
  varNameLinks(473) = Array("Q_P3_33403_2019_DF", "Plot_3_Q33403")
  varNameLinks(474) = Array("Q_ND_10_2020_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(475) = Array("Q_ND_10_2020_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks(476) = Array("Q_ND_11_2020_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(477) = Array("Q_ND_11_2020_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks(478) = Array("Q_ND_12_2020_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(479) = Array("Q_ND_12_2020_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks(480) = Array("Q_ND_13_2020_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(481) = Array("Q_ND_13_2020_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks(482) = Array("Q_ND_14_2020_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(483) = Array("Q_ND_14_2020_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks(484) = Array("Q_ND_15_2020_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(485) = Array("Q_ND_15_2020_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks(486) = Array("Q_ND_16_2020_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(487) = Array("Q_ND_16_2020_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks(488) = Array("Q_ND_17_2020_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(489) = Array("Q_ND_17_2020_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks(490) = Array("Q_ND_18_2020_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(491) = Array("Q_ND_18_2020_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks(492) = Array("Q_ND_19_2020_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(493) = Array("Q_ND_19_2020_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks(494) = Array("Q_ND_1_2020_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(495) = Array("Q_ND_1_2020_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks(496) = Array("Q_ND_20_2020_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(497) = Array("Q_ND_20_2020_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks(498) = Array("Q_ND_21_2020_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(499) = Array("Q_ND_21_2020_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks(500) = Array("Q_ND_22_2020_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(501) = Array("Q_ND_22_2020_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks(502) = Array("Q_ND_23_2020_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(503) = Array("Q_ND_23_2020_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks(504) = Array("Q_ND_24_2020_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(505) = Array("Q_ND_24_2020_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks(506) = Array("Q_ND_2_2020_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(507) = Array("Q_ND_2_2020_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks(508) = Array("Q_ND_3_2020_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(509) = Array("Q_ND_3_2020_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks(510) = Array("Q_ND_4_2020_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(511) = Array("Q_ND_4_2020_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks(512) = Array("Q_ND_5_2020_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(513) = Array("Q_ND_5_2020_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks(514) = Array("Q_ND_6_2020_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(515) = Array("Q_ND_6_2020_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks(516) = Array("Q_ND_7_2020_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(517) = Array("Q_ND_7_2020_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks(518) = Array("Q_ND_8_2020_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(519) = Array("Q_ND_8_2020_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks(520) = Array("Q_ND_9_2020_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(521) = Array("Q_ND_9_2020_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks(522) = Array("Q_Plot_1_11969_2020_CF", "Plot_1_Q11969")
  varNameLinks(523) = Array("Q_Plot_1_11969_2020_DF", "Plot_1_Q11969")
  varNameLinks(524) = Array("Q_Plot_1_387_2020_CF", "Plot_1_Q387")
  varNameLinks(525) = Array("Q_Plot_1_387_2020_DF", "Plot_1_Q387")
  varNameLinks(526) = Array("Q_Plot_1_390_2020_CF", "Plot_1_Q390")
  varNameLinks(527) = Array("Q_Plot_1_390_2020_DF", "Plot_1_Q390")
  varNameLinks(528) = Array("Q_Plot_3_11974_2020_CF", "Plot_3_Q11974")
  varNameLinks(529) = Array("Q_Plot_3_11974_2020_DF", "Plot_3_Q11974")
  varNameLinks(530) = Array("Q_Plot_3_11975_2020_CF", "Plot_3_Q11975")
  varNameLinks(531) = Array("Q_Plot_3_11975_2020_DF", "Plot_3_Q11975")
  varNameLinks(532) = Array("Q_Plot_3_11977_2020_CF", "Plot_3_Q11977")
  varNameLinks(533) = Array("Q_Plot_3_11977_2020_DF", "Plot_3_Q11977")
  varNameLinks(534) = Array("Q_Plot_3_33401_2020_CF", "Plot_3_Q33401")
  varNameLinks(535) = Array("Q_Plot_3_33401_2020_DF", "Plot_3_Q33401")
  varNameLinks(536) = Array("Q_Plot_3_33402_2020_CF", "Plot_3_Q33402")
  varNameLinks(537) = Array("Q_Plot_3_33402_2020_DF", "Plot_3_Q33402")
  varNameLinks(538) = Array("Q_Plot_3_33403_2020_CF", "Plot_3_Q33403")
  varNameLinks(539) = Array("Q_Plot_3_33403_2020_DF", "Plot_3_Q33403")
  varNameLinks(540) = Array("Q33403_2017_C", "Plot_3_Q33403")
  varNameLinks(541) = Array("Q33403_2017_D", "Plot_3_Q33403")

  AddRestOfConversions varNameLinks

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2020 = New Collection
  Set p2020toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks)
    varSubArray = varNameLinks(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2020, CStr(varSubArray(1))) Then
      pOldTo2020.Add varSubArray(0), varSubArray(1)
    End If
    p2020toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub AddRestOfConversions(varNameLinks() As Variant)

  MyGeneralOperations.AddValueToVariantArray Array("Q4_1955_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks

  MyGeneralOperations.AddValueToVariantArray Array("Q10_1935_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1935_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1935_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1935_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1935_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1935_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1935_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1935_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1935_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1935_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1935_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1935_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1935_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1935_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1935_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1935_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1935_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1935_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1935_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1935_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1935_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1935_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1935_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1935_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1935_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1935_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1935_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1935_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1935_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1935_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1935_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1935_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1935_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1935_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1935_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1935_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1935_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1935_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1395_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1935_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1935_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1935_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1935_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1935_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1935_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1935_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1935_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1935_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1936_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1936_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1936_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1936_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1936_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1936_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1936_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1936_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1936_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1936_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1936_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1936_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1936_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1936_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1936_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1936_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1936_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1936_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1936_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1936_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1936_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1936_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1936_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1936_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1936_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1936_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1936_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1936_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1936_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1936_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1936_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1936_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1936_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1936_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1936_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1936_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1936_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1936_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1936_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1936_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1936_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1936_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1936_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1936_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1936_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1936_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1936_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1936_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1940_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1940_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1940_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1940_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1940_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1940_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1940_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1940_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1940_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1940_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1940_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1940_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1940_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1940_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1940_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1940_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1940_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1940_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1940_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1940_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1940_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1940_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1940_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1940_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1940_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1940_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1940_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1940_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1940_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1940_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1940_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1940_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1940_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1940_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1940_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1940_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1940_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1940_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1940_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1940_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1940_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1940_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1940_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1940_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1941_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1941_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1941_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1941_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1942_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1942_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1942_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1942_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1942_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1942_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1942_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1942_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1942_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1942_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1942_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1942_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1942_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1942_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1942_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1942_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1942_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1942_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1942_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1942_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1942_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1942_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1942_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1942_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1942_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1942_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1942_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1942_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1942_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1942_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1942_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1942_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1942_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1942_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1942_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1942_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1942_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1942_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1942_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1942_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1942_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1942_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1942_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1942_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1942_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1942_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1942_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1942_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1952_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1952_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1952_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1952_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1952_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1952_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1952_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1952_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1952_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1952_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1952_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1952_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1952_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1952_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1952_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1952_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1952_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1952_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1952_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1952_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1952_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1952_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1952_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1952_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1952_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1952_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1952_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1952_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1952_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1952_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1952_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1952_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1952_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1952_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1952_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1952_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1952_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1952_DF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1952_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1952_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1952_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1952_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1952_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1952_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1952_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1952_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1952_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1952_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1955_CF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q10_1955_DF", "Natural_Drainages_Watershed_B_Q10"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1955_CF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11_1955_DF", "Natural_Drainages_Watershed_C_Q11"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1955_CF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q12_1955_DF", "Natural_Drainages_Watershed_D_Q12"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1955_CF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q13_1955_DF", "Natural_Drainages_Watershed_A_Q13"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1955_CF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q14_1955_DF", "Natural_Drainages_Watershed_B_Q14"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1955_CF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q15_1955_DF", "Natural_Drainages_Watershed_C_Q15"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1955_CF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q16_1955_DF", "Natural_Drainages_Watershed_D_Q16"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1955_CF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q17_1955_DF", "Natural_Drainages_Watershed_A_Q17"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1955_CF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q18_1955_DF", "Natural_Drainages_Watershed_B_Q18"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1955_CF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q19_1955_DF", "Natural_Drainages_Watershed_C_Q19"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1955_CF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q1_1955_DF", "Natural_Drainages_Watershed_A_Q1"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1955_CF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q20_1955_DF", "Natural_Drainages_Watershed_D_Q20"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1955_CF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q21_1955_DF", "Natural_Drainages_Watershed_A_Q21"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1955_CF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q22_1955_DF", "Natural_Drainages_Watershed_B_Q22"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1955_CF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q23_1955_DF", "Natural_Drainages_Watershed_C_Q23"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1955_CF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q24_1955_DF", "Natural_Drainages_Watershed_D_Q24"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1955_CF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q2_1955_DF", "Natural_Drainages_Watershed_B_Q2"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1955_CF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q3_1955_DF", "Natural_Drainages_Watershed_C_Q3"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q4_1955_CF", "Natural_Drainages_Watershed_D_Q4"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1955_CF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q5_1955_DF", "Natural_Drainages_Watershed_A_Q5"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1955_CF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q6_1955_DF", "Natural_Drainages_Watershed_B_Q6"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1955_CF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q7_1955_DF", "Natural_Drainages_Watershed_C_Q7"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1955_CF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q8_1955_DF", "Natural_Drainages_Watershed_D_Q8"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1955_CF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q9_1955_DF", "Natural_Drainages_Watershed_A_Q9"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11974_2017_CF", "Plot_3_Q11974"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11974_2017_DF", "Plot_3_Q11974"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11975_2017_CF", "Plot_3_Q11975"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11975_2017_DF", "Plot_3_Q11975"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11977_2017_CF", "Plot_3_Q11977"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q11977_2017_DF", "Plot_3_Q11977"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33401_2017_CF", "Plot_3_Q33401"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33401_2017_DF", "Plot_3_Q33401"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33402_2017_CF", "Plot_3_Q33402"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33402_2017_DF", "Plot_3_Q33402"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33404_2017_CF", "Plot_3_Q33404"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33404_2017_DF", "Plot_3_Q33404"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33401_2018_CF", "Plot_3_Q33401"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33401_2018_DF", "Plot_3_Q33401"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33402_2018_CF", "Plot_3_Q33402"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33402_2018_DF", "Plot_3_Q33402"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33403_2018_CF", "Plot_3_Q33403"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("QP3_33403_2018_DF", "Plot_3_Q33403"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33403_2017_CF", "Plot_3_Q33403"), varNameLinks
  MyGeneralOperations.AddValueToVariantArray Array("Q33403_2017_DF", "Plot_3_Q33403"), varNameLinks

End Sub

Public Sub FillNameConverters_2024(varNameLinks_2024() As Variant, p2024toOld As Collection, pOldTo2024 As Collection)

  ReDim varNameLinks_2024(93)

  varNameLinks_2024(0) = Array("Q_ND_10_2024_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2024(1) = Array("Q_ND_10_2024_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2024(2) = Array("Q_ND_11_2024_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2024(3) = Array("Q_ND_11_2024_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2024(4) = Array("Q_ND_12_2024_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2024(5) = Array("Q_ND_12_2024_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2024(6) = Array("Q_ND_13_2024_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2024(7) = Array("Q_ND_13_2024_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2024(8) = Array("Q_ND_14_2024_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2024(9) = Array("Q_ND_14_2024_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2024(10) = Array("Q_ND_15_2024_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2024(11) = Array("Q_ND_15_2024_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2024(12) = Array("Q_ND_16_2024_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2024(13) = Array("Q_ND_16_2024_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2024(14) = Array("Q_ND_17_2024_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2024(15) = Array("Q_ND_17_2024_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2024(16) = Array("Q_ND_18_2024_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2024(17) = Array("Q_ND_18_2024_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2024(18) = Array("Q_ND_19_2024_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2024(19) = Array("Q_ND_19_2024_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2024(20) = Array("Q_ND_1_2024_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2024(21) = Array("Q_ND_1_2024_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2024(22) = Array("Q_ND_20_2024_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2024(23) = Array("Q_ND_20_2024_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2024(24) = Array("Q_ND_21_2024_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2024(25) = Array("Q_ND_21_2024_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2024(26) = Array("Q_ND_22_2024_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2024(27) = Array("Q_ND_22_2024_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2024(28) = Array("Q_ND_23_2024_CF_NO_sp", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2024(29) = Array("Q_ND_23_2024_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2024(30) = Array("Q_ND_24_2024_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2024(31) = Array("Q_ND_24_2024_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2024(32) = Array("Q_ND_2_2024_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2024(33) = Array("Q_ND_2_2024_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2024(34) = Array("Q_ND_3_2024_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2024(35) = Array("Q_ND_3_2024_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2024(36) = Array("Q_ND_4_2024_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2024(37) = Array("Q_ND_4_2024_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2024(38) = Array("Q_ND_5_2024_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2024(39) = Array("Q_ND_5_2024_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2024(40) = Array("Q_ND_6_2024_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2024(41) = Array("Q_ND_6_2024_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2024(42) = Array("Q_ND_7_2024_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2024(43) = Array("Q_ND_7_2024_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2024(44) = Array("Q_ND_8_2024_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2024(45) = Array("Q_ND_8_2024_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2024(46) = Array("Q_ND_9_2024_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2024(47) = Array("Q_ND_9_2024_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2024(48) = Array("Q_P1_11969_2024_CF", "Plot_1_Q11969")
  varNameLinks_2024(49) = Array("Q_P1_11969_2024_DF", "Plot_1_Q11969")
  varNameLinks_2024(50) = Array("Q_P1_11972_2024_CF", "Plot_1_Q11972")
  varNameLinks_2024(51) = Array("Q_P1_11972_2024_DF", "Plot_1_Q11972")
  varNameLinks_2024(52) = Array("Q_P1_387_2024_CF", "Plot_1_Q387")
  varNameLinks_2024(53) = Array("Q_P1_387_2024_DF", "Plot_1_Q387")
  varNameLinks_2024(54) = Array("Q_P1_390_2024_CF", "Plot_1_Q390")
  varNameLinks_2024(55) = Array("Q_P1_390_2024_DF", "Plot_1_Q390")
  varNameLinks_2024(56) = Array("Q_P1_39981_2024_CF", "Plot_1_Q39981")
  varNameLinks_2024(57) = Array("Q_P1_39981_2024_DF", "Plot_1_Q39981")
  varNameLinks_2024(58) = Array("Q_P2_11962_2024_CF", "Plot_2_Q11962")
  varNameLinks_2024(59) = Array("Q_P2_11962_2024_DF", "Plot_2_Q11962")
  varNameLinks_2024(60) = Array("Q_P2_11965_2024_CF", "Plot_2_Q11965")
  varNameLinks_2024(61) = Array("Q_P2_11965_2024_DF", "Plot_2_Q11965")
  varNameLinks_2024(62) = Array("Q_P2_33410_2024_CF", "Plot_2_Q33410")
  varNameLinks_2024(63) = Array("Q_P2_33410_2024_DF", "Plot_2_Q33410")
  varNameLinks_2024(64) = Array("Q_P3_11974_2024_CF", "Plot_3_Q11974")
  varNameLinks_2024(65) = Array("Q_P3_11974_2024_DF", "Plot_3_Q11974")
  varNameLinks_2024(66) = Array("Q_P3_11975_2024_CF", "Plot_3_Q11975")
  varNameLinks_2024(67) = Array("Q_P3_11975_2024_DF", "Plot_3_Q11975")
  varNameLinks_2024(68) = Array("Q_P3_11977_2024_CF", "Plot_3_Q11977")
  varNameLinks_2024(69) = Array("Q_P3_11977_2024_DF", "Plot_3_Q11977")
  varNameLinks_2024(70) = Array("Q_P3_33401_2024_CF", "Plot_3_Q33401")
  varNameLinks_2024(71) = Array("Q_P3_33401_2024_DF", "Plot_3_Q33401")
  varNameLinks_2024(72) = Array("Q_P3_33402_2024_CF", "Plot_3_Q33402")
  varNameLinks_2024(73) = Array("Q_P3_33402_2024_DF", "Plot_3_Q33402")
  varNameLinks_2024(74) = Array("Q_P3_33403_2024_CF", "Plot_3_Q33403")
  varNameLinks_2024(75) = Array("Q_P3_33403_2024_DF", "Plot_3_Q33403")
  varNameLinks_2024(76) = Array("Q_S_26358_2024_CF", "Summit_Plots_Q26358")
  varNameLinks_2024(77) = Array("Q_S_26358_2024_DF", "Summit_Plots_Q26358")
  varNameLinks_2024(78) = Array("Q_S_26361_2024_CF", "Summit_Plots_Q26361")
  varNameLinks_2024(79) = Array("Q_S_26361_2024_DF", "Summit_Plots_Q26361")
  varNameLinks_2024(80) = Array("Q_S_27302_2024_CF", "Summit_Plots_Q27302")
  varNameLinks_2024(81) = Array("Q_S_27302_2024_DF", "Summit_Plots_Q27302")
  varNameLinks_2024(82) = Array("Q_S_27313_2024_CF", "Summit_Plots_Q27313")
  varNameLinks_2024(83) = Array("Q_S_27313_2024_DF", "Summit_Plots_Q27313")
  varNameLinks_2024(84) = Array("Q_S_27332_2024_CF", "Summit_Plots_Q27332")
  varNameLinks_2024(85) = Array("Q_S_27332_2024_DF", "Summit_Plots_Q27332")
  varNameLinks_2024(86) = Array("Q_S_27335_2024_CF", "Summit_Plots_Q27335")
  varNameLinks_2024(87) = Array("Q_S_27335_2024_DF", "Summit_Plots_Q27335")
  varNameLinks_2024(88) = Array("Q_S_29937_2024_CF", "Summit_Plots_Q29937")
  varNameLinks_2024(89) = Array("Q_S_29937_2024_DF", "Summit_Plots_Q29937")
  varNameLinks_2024(90) = Array("Q_S_393_2024_CF", "Summit_Plots_Q393")
  varNameLinks_2024(91) = Array("Q_S_393_2024_DF", "Summit_Plots_Q393")
  varNameLinks_2024(92) = Array("Q_S_54_2024_CF", "Summit_Plots_Q54")
  varNameLinks_2024(93) = Array("Q_S_54_2024_DF", "Summit_Plots_Q54")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2024 = New Collection
  Set p2024toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2024)
    varSubArray = varNameLinks_2024(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2024, CStr(varSubArray(1))) Then
      pOldTo2024.Add varSubArray(0), varSubArray(1)
    End If
    p2024toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2023(varNameLinks_2023() As Variant, p2023toOld As Collection, pOldTo2023 As Collection)

  ReDim varNameLinks_2023(93)

  varNameLinks_2023(0) = Array("Q_ND_10_2023_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2023(1) = Array("Q_ND_10_2023_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2023(2) = Array("Q_ND_11_2023_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2023(3) = Array("Q_ND_11_2023_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2023(4) = Array("Q_ND_12_2023_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2023(5) = Array("Q_ND_12_2023_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2023(6) = Array("Q_ND_13_2023_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2023(7) = Array("Q_ND_13_2023_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2023(8) = Array("Q_ND_14_2023_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2023(9) = Array("Q_ND_14_2023_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2023(10) = Array("Q_ND_15_2023_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2023(11) = Array("Q_ND_15_2023_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2023(12) = Array("Q_ND_16_2023_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2023(13) = Array("Q_ND_16_2023_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2023(14) = Array("Q_ND_17_2023_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2023(15) = Array("Q_ND_17_2023_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2023(16) = Array("Q_ND_18_2023_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2023(17) = Array("Q_ND_18_2023_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2023(18) = Array("Q_ND_19_2023_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2023(19) = Array("Q_ND_19_2023_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2023(20) = Array("Q_ND_1_2023_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2023(21) = Array("Q_ND_1_2023_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2023(22) = Array("Q_ND_20_2023_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2023(23) = Array("Q_ND_20_2023_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2023(24) = Array("Q_ND_21_2023_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2023(25) = Array("Q_ND_21_2023_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2023(26) = Array("Q_ND_22_2023_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2023(27) = Array("Q_ND_22_2023_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2023(28) = Array("Q_ND_23_2023_CF_NO_sp", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2023(29) = Array("Q_ND_23_2023_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2023(30) = Array("Q_ND_24_2023_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2023(31) = Array("Q_ND_24_2023_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2023(32) = Array("Q_ND_2_2023_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2023(33) = Array("Q_ND_2_2023_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2023(34) = Array("Q_ND_3_2023_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2023(35) = Array("Q_ND_3_2023_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2023(36) = Array("Q_ND_4_2023_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2023(37) = Array("Q_ND_4_2023_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2023(38) = Array("Q_ND_5_2023_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2023(39) = Array("Q_ND_5_2023_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2023(40) = Array("Q_ND_6_2023_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2023(41) = Array("Q_ND_6_2023_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2023(42) = Array("Q_ND_7_2023_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2023(43) = Array("Q_ND_7_2023_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2023(44) = Array("Q_ND_8_2023_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2023(45) = Array("Q_ND_8_2023_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2023(46) = Array("Q_ND_9_2023_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2023(47) = Array("Q_ND_9_2023_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2023(48) = Array("Q_P1_11969_2023_CF", "Plot_1_Q11969")
  varNameLinks_2023(49) = Array("Q_P1_11969_2023_DF", "Plot_1_Q11969")
  varNameLinks_2023(50) = Array("Q_P1_11972_2023_CF", "Plot_1_Q11972")
  varNameLinks_2023(51) = Array("Q_P1_11972_2023_DF", "Plot_1_Q11972")
  varNameLinks_2023(52) = Array("Q_P1_387_2023_CF", "Plot_1_Q387")
  varNameLinks_2023(53) = Array("Q_P1_387_2023_DF", "Plot_1_Q387")
  varNameLinks_2023(54) = Array("Q_P1_390_2023_CF", "Plot_1_Q390")
  varNameLinks_2023(55) = Array("Q_P1_390_2023_DF", "Plot_1_Q390")
  varNameLinks_2023(56) = Array("Q_P1_39981_2023_CF", "Plot_1_Q39981")
  varNameLinks_2023(57) = Array("Q_P1_39981_2023_DF", "Plot_1_Q39981")
  varNameLinks_2023(58) = Array("Q_P2_11962_2023_CF", "Plot_2_Q11962")
  varNameLinks_2023(59) = Array("Q_P2_11962_2023_DF", "Plot_2_Q11962")
  varNameLinks_2023(60) = Array("Q_P2_11965_2023_CF", "Plot_2_Q11965")
  varNameLinks_2023(61) = Array("Q_P2_11965_2023_DF", "Plot_2_Q11965")
  varNameLinks_2023(62) = Array("Q_P2_33410_2023_CF", "Plot_2_Q33410")
  varNameLinks_2023(63) = Array("Q_P2_33410_2023_DF", "Plot_2_Q33410")
  varNameLinks_2023(64) = Array("Q_P3_11974_2023_CF", "Plot_3_Q11974")
  varNameLinks_2023(65) = Array("Q_P3_11974_2023_DF", "Plot_3_Q11974")
  varNameLinks_2023(66) = Array("Q_P3_11975_2023_CF", "Plot_3_Q11975")
  varNameLinks_2023(67) = Array("Q_P3_11975_2023_DF", "Plot_3_Q11975")
  varNameLinks_2023(68) = Array("Q_P3_11977_2023_CF", "Plot_3_Q11977")
  varNameLinks_2023(69) = Array("Q_P3_11977_2023_DF", "Plot_3_Q11977")
  varNameLinks_2023(70) = Array("Q_P3_33401_2023_CF", "Plot_3_Q33401")
  varNameLinks_2023(71) = Array("Q_P3_33401_2023_DF", "Plot_3_Q33401")
  varNameLinks_2023(72) = Array("Q_P3_33402_2023_CF", "Plot_3_Q33402")
  varNameLinks_2023(73) = Array("Q_P3_33402_2023_DF", "Plot_3_Q33402")
  varNameLinks_2023(74) = Array("Q_P3_33403_2023_CF", "Plot_3_Q33403")
  varNameLinks_2023(75) = Array("Q_P3_33403_2023_DF", "Plot_3_Q33403")
  varNameLinks_2023(76) = Array("Q_S_26358_2023_CF", "Summit_Plots_Q26358")
  varNameLinks_2023(77) = Array("Q_S_26358_2023_DF", "Summit_Plots_Q26358")
  varNameLinks_2023(78) = Array("Q_S_26361_2023_CF", "Summit_Plots_Q26361")
  varNameLinks_2023(79) = Array("Q_S_26361_2023_DF", "Summit_Plots_Q26361")
  varNameLinks_2023(80) = Array("Q_S_27302_2023_CF", "Summit_Plots_Q27302")
  varNameLinks_2023(81) = Array("Q_S_27302_2023_DF", "Summit_Plots_Q27302")
  varNameLinks_2023(82) = Array("Q_S_27313_2023_CF", "Summit_Plots_Q27313")
  varNameLinks_2023(83) = Array("Q_S_27313_2023_DF", "Summit_Plots_Q27313")
  varNameLinks_2023(84) = Array("Q_S_27332_2023_CF", "Summit_Plots_Q27332")
  varNameLinks_2023(85) = Array("Q_S_27332_2023_DF", "Summit_Plots_Q27332")
  varNameLinks_2023(86) = Array("Q_S_27335_2023_CF", "Summit_Plots_Q27335")
  varNameLinks_2023(87) = Array("Q_S_27335_2023_DF", "Summit_Plots_Q27335")
  varNameLinks_2023(88) = Array("Q_S_29937_2023_CF", "Summit_Plots_Q29937")
  varNameLinks_2023(89) = Array("Q_S_29937_2023_DF", "Summit_Plots_Q29937")
  varNameLinks_2023(90) = Array("Q_S_393_2023_CF", "Summit_Plots_Q393")
  varNameLinks_2023(91) = Array("Q_S_393_2023_DF", "Summit_Plots_Q393")
  varNameLinks_2023(92) = Array("Q_S_54_2023_CF", "Summit_Plots_Q54")
  varNameLinks_2023(93) = Array("Q_S_54_2023_DF", "Summit_Plots_Q54")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2023 = New Collection
  Set p2023toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2023)
    varSubArray = varNameLinks_2023(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2023, CStr(varSubArray(1))) Then
      pOldTo2023.Add varSubArray(0), varSubArray(1)
    End If
    p2023toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2022(varNameLinks_2022() As Variant, p2022toOld As Collection, pOldTo2022 As Collection)

  ReDim varNameLinks_2022(93)

  varNameLinks_2022(0) = Array("Q_ND_10_2022_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2022(1) = Array("Q_ND_10_2022_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2022(2) = Array("Q_ND_11_2022_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2022(3) = Array("Q_ND_11_2022_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2022(4) = Array("Q_ND_12_2022_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2022(5) = Array("Q_ND_12_2022_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2022(6) = Array("Q_ND_13_2022_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2022(7) = Array("Q_ND_13_2022_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2022(8) = Array("Q_ND_14_2022_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2022(9) = Array("Q_ND_14_2022_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2022(10) = Array("Q_ND_15_2022_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2022(11) = Array("Q_ND_15_2022_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2022(12) = Array("Q_ND_16_2022_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2022(13) = Array("Q_ND_16_2022_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2022(14) = Array("Q_ND_17_2022_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2022(15) = Array("Q_ND_17_2022_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2022(16) = Array("Q_ND_18_2022_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2022(17) = Array("Q_ND_18_2022_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2022(18) = Array("Q_ND_19_2022_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2022(19) = Array("Q_ND_19_2022_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2022(20) = Array("Q_ND_1_2022_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2022(21) = Array("Q_ND_1_2022_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2022(22) = Array("Q_ND_20_2022_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2022(23) = Array("Q_ND_20_2022_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2022(24) = Array("Q_ND_21_2022_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2022(25) = Array("Q_ND_21_2022_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2022(26) = Array("Q_ND_22_2022_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2022(27) = Array("Q_ND_22_2022_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2022(28) = Array("Q_ND_23_2022_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2022(29) = Array("Q_ND_23_2022_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2022(30) = Array("Q_ND_24_2022_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2022(31) = Array("Q_ND_24_2022_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2022(32) = Array("Q_ND_2_2022_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2022(33) = Array("Q_ND_2_2022_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2022(34) = Array("Q_ND_3_2022_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2022(35) = Array("Q_ND_3_2022_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2022(36) = Array("Q_ND_4_2022_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2022(37) = Array("Q_ND_4_2022_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2022(38) = Array("Q_ND_5_2022_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2022(39) = Array("Q_ND_5_2022_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2022(40) = Array("Q_ND_6_2022_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2022(41) = Array("Q_ND_6_2022_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2022(42) = Array("Q_ND_7_2022_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2022(43) = Array("Q_ND_7_2022_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2022(44) = Array("Q_ND_8_2022_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2022(45) = Array("Q_ND_8_2022_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2022(46) = Array("Q_ND_9_2022_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2022(47) = Array("Q_ND_9_2022_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2022(48) = Array("Q_P1_11969_2022_CF", "Plot_1_Q11969")
  varNameLinks_2022(49) = Array("Q_P1_11969_2022_DF", "Plot_1_Q11969")
  varNameLinks_2022(50) = Array("Q_P1_11972_2022_CF", "Plot_1_Q11972")
  varNameLinks_2022(51) = Array("Q_P1_11972_2022_DF", "Plot_1_Q11972")
  varNameLinks_2022(52) = Array("Q_P1_387_2022_CF", "Plot_1_Q387")
  varNameLinks_2022(53) = Array("Q_P1_387_2022_DF", "Plot_1_Q387")
  varNameLinks_2022(54) = Array("Q_P1_390_2022_CF", "Plot_1_Q390")
  varNameLinks_2022(55) = Array("Q_P1_390_2022_DF", "Plot_1_Q390")
  varNameLinks_2022(56) = Array("Q_P1_39981_2022_CF", "Plot_1_Q39981")
  varNameLinks_2022(57) = Array("Q_P1_39981_2022_DF", "Plot_1_Q39981")
  varNameLinks_2022(58) = Array("Q_P2_11962_2022_CF", "Plot_2_Q11962")
  varNameLinks_2022(59) = Array("Q_P2_11962_2022_DF", "Plot_2_Q11962")
  varNameLinks_2022(60) = Array("Q_P2_11965_2022_CF", "Plot_2_Q11965")
  varNameLinks_2022(61) = Array("Q_P2_11965_2022_DF", "Plot_2_Q11965")
  varNameLinks_2022(62) = Array("Q_P2_33410_2022_CF", "Plot_2_Q33410")
  varNameLinks_2022(63) = Array("Q_P2_33410_2022_DF", "Plot_2_Q33410")
  varNameLinks_2022(64) = Array("Q_P3_11974_2022_CF", "Plot_3_Q11974")
  varNameLinks_2022(65) = Array("Q_P3_11974_2022_DF", "Plot_3_Q11974")
  varNameLinks_2022(66) = Array("Q_P3_11975_2022_CF", "Plot_3_Q11975")
  varNameLinks_2022(67) = Array("Q_P3_11975_2022_DF", "Plot_3_Q11975")
  varNameLinks_2022(68) = Array("Q_P3_11977_2022_CF", "Plot_3_Q11977")
  varNameLinks_2022(69) = Array("Q_P3_11977_2022_DF", "Plot_3_Q11977")
  varNameLinks_2022(70) = Array("Q_P3_33401_2022_CF", "Plot_3_Q33401")
  varNameLinks_2022(71) = Array("Q_P3_33401_2022_DF", "Plot_3_Q33401")
  varNameLinks_2022(72) = Array("Q_P3_33402_2022_CF", "Plot_3_Q33402")
  varNameLinks_2022(73) = Array("Q_P3_33402_2022_DF", "Plot_3_Q33402")
  varNameLinks_2022(74) = Array("Q_P3_33403_2022_CF", "Plot_3_Q33403")
  varNameLinks_2022(75) = Array("Q_P3_33403_2022_DF", "Plot_3_Q33403")
  varNameLinks_2022(76) = Array("Q_S_26358_2022_CF", "Summit_Plots_Q26358")
  varNameLinks_2022(77) = Array("Q_S_26358_2022_DF", "Summit_Plots_Q26358")
  varNameLinks_2022(78) = Array("Q_S_26361_2022_CF", "Summit_Plots_Q26361")
  varNameLinks_2022(79) = Array("Q_S_26361_2022_DF", "Summit_Plots_Q26361")
  varNameLinks_2022(80) = Array("Q_S_27302_2022_CF", "Summit_Plots_Q27302")
  varNameLinks_2022(81) = Array("Q_S_27302_2022_DF", "Summit_Plots_Q27302")
  varNameLinks_2022(82) = Array("Q_S_27313_2022_CF", "Summit_Plots_Q27313")
  varNameLinks_2022(83) = Array("Q_S_27313_2022_DF", "Summit_Plots_Q27313")
  varNameLinks_2022(84) = Array("Q_S_27332_2022_CF", "Summit_Plots_Q27332")
  varNameLinks_2022(85) = Array("Q_S_27332_2022_DF", "Summit_Plots_Q27332")
  varNameLinks_2022(86) = Array("Q_S_27335_2022_CF", "Summit_Plots_Q27335")
  varNameLinks_2022(87) = Array("Q_S_27335_2022_DF", "Summit_Plots_Q27335")
  varNameLinks_2022(88) = Array("Q_S_29937_2022_CF", "Summit_Plots_Q29937")
  varNameLinks_2022(89) = Array("Q_S_29937_2022_DF", "Summit_Plots_Q29937")
  varNameLinks_2022(90) = Array("Q_S_393_2022_CF", "Summit_Plots_Q393")
  varNameLinks_2022(91) = Array("Q_S_393_2022_DF", "Summit_Plots_Q393")
  varNameLinks_2022(92) = Array("Q_S_54_2022_CF", "Summit_Plots_Q54")
  varNameLinks_2022(93) = Array("Q_S_54_2022_DF", "Summit_Plots_Q54")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2022 = New Collection
  Set p2022toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2022)
    varSubArray = varNameLinks_2022(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2022, CStr(varSubArray(1))) Then
      pOldTo2022.Add varSubArray(0), varSubArray(1)
    End If
    p2022toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub FillNameConverters_2021(varNameLinks_2021() As Variant, p2021toOld As Collection, pOldTo2021 As Collection)

  ReDim varNameLinks_2021(87)

  varNameLinks_2021(0) = Array("Q_ND_10_2021_CF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2021(1) = Array("Q_ND_10_2021_DF", "Natural_Drainages_Watershed_B_Q10")
  varNameLinks_2021(2) = Array("Q_ND_11_2021_CF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2021(3) = Array("Q_ND_11_2021_DF", "Natural_Drainages_Watershed_C_Q11")
  varNameLinks_2021(4) = Array("Q_ND_12_2021_CF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2021(5) = Array("Q_ND_12_2021_DF", "Natural_Drainages_Watershed_D_Q12")
  varNameLinks_2021(6) = Array("Q_ND_13_2021_CF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2021(7) = Array("Q_ND_13_2021_DF", "Natural_Drainages_Watershed_A_Q13")
  varNameLinks_2021(8) = Array("Q_ND_14_2021_CF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2021(9) = Array("Q_ND_14_2021_DF", "Natural_Drainages_Watershed_B_Q14")
  varNameLinks_2021(10) = Array("Q_ND_15_2021_CF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2021(11) = Array("Q_ND_15_2021_DF", "Natural_Drainages_Watershed_C_Q15")
  varNameLinks_2021(12) = Array("Q_ND_16_2021_CF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2021(13) = Array("Q_ND_16_2021_DF", "Natural_Drainages_Watershed_D_Q16")
  varNameLinks_2021(14) = Array("Q_ND_17_2021_CF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2021(15) = Array("Q_ND_17_2021_DF", "Natural_Drainages_Watershed_A_Q17")
  varNameLinks_2021(16) = Array("Q_ND_18_2021_CF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2021(17) = Array("Q_ND_18_2021_DF", "Natural_Drainages_Watershed_B_Q18")
  varNameLinks_2021(18) = Array("Q_ND_19_2021_CF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2021(19) = Array("Q_ND_19_2021_DF", "Natural_Drainages_Watershed_C_Q19")
  varNameLinks_2021(20) = Array("Q_ND_1_2021_CF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2021(21) = Array("Q_ND_1_2021_DF", "Natural_Drainages_Watershed_A_Q1")
  varNameLinks_2021(22) = Array("Q_ND_20_2021_CF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2021(23) = Array("Q_ND_20_2021_DF", "Natural_Drainages_Watershed_D_Q20")
  varNameLinks_2021(24) = Array("Q_ND_21_2021_CF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2021(25) = Array("Q_ND_21_2021_DF", "Natural_Drainages_Watershed_A_Q21")
  varNameLinks_2021(26) = Array("Q_ND_22_2021_CF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2021(27) = Array("Q_ND_22_2021_DF", "Natural_Drainages_Watershed_B_Q22")
  varNameLinks_2021(28) = Array("Q_ND_23_2021_CF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2021(29) = Array("Q_ND_23_2021_DF", "Natural_Drainages_Watershed_C_Q23")
  varNameLinks_2021(30) = Array("Q_ND_24_2021_CF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2021(31) = Array("Q_ND_24_2021_DF", "Natural_Drainages_Watershed_D_Q24")
  varNameLinks_2021(32) = Array("Q_ND_2_2021_CF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2021(33) = Array("Q_ND_2_2021_DF", "Natural_Drainages_Watershed_B_Q2")
  varNameLinks_2021(34) = Array("Q_ND_3_2021_CF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2021(35) = Array("Q_ND_3_2021_DF", "Natural_Drainages_Watershed_C_Q3")
  varNameLinks_2021(36) = Array("Q_ND_4_2021_CF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2021(37) = Array("Q_ND_4_2021_DF", "Natural_Drainages_Watershed_D_Q4")
  varNameLinks_2021(38) = Array("Q_ND_5_2021_CF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2021(39) = Array("Q_ND_5_2021_DF", "Natural_Drainages_Watershed_A_Q5")
  varNameLinks_2021(40) = Array("Q_ND_6_2021_CF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2021(41) = Array("Q_ND_6_2021_DF", "Natural_Drainages_Watershed_B_Q6")
  varNameLinks_2021(42) = Array("Q_ND_7_2021_CF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2021(43) = Array("Q_ND_7_2021_DF", "Natural_Drainages_Watershed_C_Q7")
  varNameLinks_2021(44) = Array("Q_ND_8_2021_CF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2021(45) = Array("Q_ND_8_2021_DF", "Natural_Drainages_Watershed_D_Q8")
  varNameLinks_2021(46) = Array("Q_ND_9_2021_CF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2021(47) = Array("Q_ND_9_2021_DF", "Natural_Drainages_Watershed_A_Q9")
  varNameLinks_2021(48) = Array("Q_P1_11969_2021_CF", "Plot_1_Q11969")
  varNameLinks_2021(49) = Array("Q_P1_11969_2021_DF", "Plot_1_Q11969")
  varNameLinks_2021(50) = Array("Q_P1_11972_2021_CF", "Plot_1_Q11972")
  varNameLinks_2021(51) = Array("Q_P1_11972_2021_DF", "Plot_1_Q11972")
  varNameLinks_2021(52) = Array("Q_P1_387_2021_CF", "Plot_1_Q387")
  varNameLinks_2021(53) = Array("Q_P1_387_2021_DF", "Plot_1_Q387")
  varNameLinks_2021(54) = Array("Q_P1_390_2021_CF", "Plot_1_Q390")
  varNameLinks_2021(55) = Array("Q_P1_390_2021_DF", "Plot_1_Q390")
  varNameLinks_2021(56) = Array("Q_P3_11974_2021_CF", "Plot_3_Q11974")
  varNameLinks_2021(57) = Array("Q_P3_11974_2021_DF", "Plot_3_Q11974")
  varNameLinks_2021(58) = Array("Q_P3_11975_2021_CF", "Plot_3_Q11975")
  varNameLinks_2021(59) = Array("Q_P3_11975_2021_DF", "Plot_3_Q11975")
  varNameLinks_2021(60) = Array("Q_P3_11977_2021_CF", "Plot_3_Q11977")
  varNameLinks_2021(61) = Array("Q_P3_11977_2021_DF", "Plot_3_Q11977")
  varNameLinks_2021(62) = Array("Q_P3_33401_2021_CF", "Plot_3_Q33401")
  varNameLinks_2021(63) = Array("Q_P3_33401_2021_DF", "Plot_3_Q33401")
  varNameLinks_2021(64) = Array("Q_P3_33402_2021_CF", "Plot_3_Q33402")
  varNameLinks_2021(65) = Array("Q_P3_33402_2021_DF", "Plot_3_Q33402")
  varNameLinks_2021(66) = Array("Q_P3_33403_2021_CF", "Plot_3_Q33403")
  varNameLinks_2021(67) = Array("Q_P3_33403_2021_DF", "Plot_3_Q33403")
  varNameLinks_2021(68) = Array("Q_S_26358_2021_CF", "Summit_Plots_Q26358")
  varNameLinks_2021(69) = Array("Q_S_26358_2021_DF", "Summit_Plots_Q26358")
  varNameLinks_2021(70) = Array("Q_S_26359_2021_CF", "Summit_Plots_Q26359")
  varNameLinks_2021(71) = Array("Q_S_26359_2021_DF", "Summit_Plots_Q26359")
  varNameLinks_2021(72) = Array("Q_S_26361_2021_CF", "Summit_Plots_Q26361")
  varNameLinks_2021(73) = Array("Q_S_26361_2021_DF", "Summit_Plots_Q26361")
  varNameLinks_2021(74) = Array("Q_S_27302_2021_CF", "Summit_Plots_Q27302")
  varNameLinks_2021(75) = Array("Q_S_27302_2021_DF", "Summit_Plots_Q27302")
  varNameLinks_2021(76) = Array("Q_S_27313_2021_CF", "Summit_Plots_Q27313")
  varNameLinks_2021(77) = Array("Q_S_27313_2021_DF", "Summit_Plots_Q27313")
  varNameLinks_2021(78) = Array("Q_S_27332_2021_CF", "Summit_Plots_Q27332")
  varNameLinks_2021(79) = Array("Q_S_27332_2021_DF", "Summit_Plots_Q27332")
  varNameLinks_2021(80) = Array("Q_S_27335_2021_CF", "Summit_Plots_Q27335")
  varNameLinks_2021(81) = Array("Q_S_27335_2021_DF", "Summit_Plots_Q27335")
  varNameLinks_2021(82) = Array("Q_S_29937_2021_CF", "Summit_Plots_Q29937")
  varNameLinks_2021(83) = Array("Q_S_29937_2021_DF", "Summit_Plots_Q29937")
  varNameLinks_2021(84) = Array("Q_S_391_2021_CF", "Summit_Plots_Q391")
  varNameLinks_2021(85) = Array("Q_S_391_2021_DF", "Summit_Plots_Q391")
  varNameLinks_2021(86) = Array("Q_S_393_2021_CF", "Summit_Plots_Q393")
  varNameLinks_2021(87) = Array("Q_S_393_2021_DF", "Summit_Plots_Q393")

  Dim lngIndex As Long
  Dim varSubArray() As Variant
  Set pOldTo2021 = New Collection
  Set p2021toOld = New Collection

  For lngIndex = 0 To UBound(varNameLinks_2021)
    varSubArray = varNameLinks_2021(lngIndex)
    If Not MyGeneralOperations.CheckCollectionForKey(pOldTo2021, CStr(varSubArray(1))) Then
      pOldTo2021.Add varSubArray(0), varSubArray(1)
    End If
    p2021toOld.Add varSubArray(1), varSubArray(0)
  Next lngIndex

End Sub

Public Sub ConvertPointShapefiles_SA()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pCoverCollection As New Collection
  Dim pDensityCollection As New Collection

  Dim pCoverToDensity As Collection
  Dim pDensityToCover As Collection
  Dim strCoverToDensityQuery As String
  Dim strDensityToCoverQuery As String
  Dim pCoverShouldChangeColl As Collection
  Dim pDensityShouldChangeColl As Collection
  Dim pRotateColl As Collection

  Debug.Print "---------------------"
  Call FillCollections_SA(pCoverCollection, pDensityCollection, pCoverToDensity, pDensityToCover, _
    strCoverToDensityQuery, strDensityToCoverQuery, pCoverShouldChangeColl, pDensityShouldChangeColl)

  Set pRotateColl = FillRotateColl_SA
  Dim pRotator As ITransform2D
  Dim strRotateBy As String
  Dim dblRotateVal As Double
  Dim pCollByQuadrat As Collection
  Dim varRotateElements() As Variant
  Dim pMidPoint As IPoint
  Set pMidPoint = New Point
  pMidPoint.PutCoords 0.5, 0.5

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String

  Dim strNewRoot As String
  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, strNewRoot, , , , strContainingFolder)

  Dim strMissingSpeciesPath As String
  Dim lngFileNum As Long
  Dim strSpeciesArray() As String
  Dim strMissingSummaryPath As String
  strMissingSpeciesPath = strContainingFolder & "\Missing_Species.csv"
  strMissingSummaryPath = strContainingFolder & "\Missing_Species_Summary.csv"

  lngFileNum = FreeFile(0)
  Open strMissingSpeciesPath For Output As lngFileNum
  Print #lngFileNum, """Quadrat"",""Species"""
  Close #lngFileNum

  lngFileNum = FreeFile(0)
  Open strMissingSummaryPath For Output As lngFileNum
  Print #lngFileNum, """Species"",""Quadrats"""
  Close #lngFileNum

  Dim pSpeciesSummaryColl As New Collection
  Dim pSubColl As Collection
  Dim strSubNames() As String
  Dim varSubArray() As Variant
  Dim strSpeciesLine As String

  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long
  Dim lngIndex2 As Long

  Dim pDataset As IDataset
  Dim booFoundShapefiles As Boolean
  Dim varDatasets() As Variant

  Dim strNames() As String
  Dim strName As String
  Dim lngDatasetIndex As Long
  Dim lngNameIndex As Long
  Dim lngNameCount As Long
  Dim booFoundNames As Boolean
  Dim lngRecCount As Long

  Dim strFullNames() As String
  Dim lngFullNameCounter As Long

  Dim lngShapefileCount As Long
  Dim lngAcceptSFCount As Long
  lngShapefileCount = 0
  lngRecCount = 0
  lngAcceptSFCount = 0

  lngFullNameCounter = -1
  Dim pNameColl As New Collection
  Dim strHexify As String
  Dim strCorrect As String
  Dim pCheckCollection As Collection
  Dim strReport As String
  Dim booMadeChanges As Boolean
  Dim strEditReport As String
  Dim strExcelReport As String
  Dim strExcelFullReport As String
  Dim pFClass As IFeatureClass
  Dim strBase As String

  Dim strFolderName As String
  Dim booFoundPolys As Boolean
  Dim booFoundPoints As Boolean
  Dim pRepPointFClass As IFeatureClass
  Dim pRepPolyFClass As IFeatureClass
  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory
  Dim pField As iField
  Dim pNewFields As esriSystem.IVariantArray

  Dim pNewDensityFClass As IFeatureClass
  Dim varDensityFieldIndexArray() As Variant
  Dim strNewDensityFClassName As String
  Dim booDensityHasFields As Boolean
  Dim lngDensityFClassIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensityYearIndex As Long
  Dim lngDensityTypeIndex As Long
  Dim lngDensityOrigFIDIndex As Long
  Dim lngDensityRotationIndex As Long

  Dim pNewGDBDensityFClass As IFeatureClass
  Dim varGDBDensityFieldIndexArray() As Variant
  Dim strGDBNewDensityFClassName As String
  Dim booGDBDensityHasFields As Boolean
  Dim lngGDBDensityFClassIndex As Long
  Dim lngGDBDensityQuadratIndex As Long
  Dim lngGDBDensityYearIndex As Long
  Dim lngGDBDensityTypeIndex As Long
  Dim lngGDBDensityOrigFIDIndex As Long
  Dim lngGDBDensityRotationIndex As Long

  Dim pNewCoverFClass As IFeatureClass
  Dim varCoverFieldIndexArray() As Variant
  Dim strNewCoverFClassName As String
  Dim booCoverHasFields As Boolean
  Dim lngCoverFClassIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverTypeIndex As Long
  Dim lngCoverOrigFIDIndex As Long
  Dim lngCoverRotationIndex As Long

  Dim pNewGDBCoverFClass As IFeatureClass
  Dim varGDBCoverFieldIndexArray() As Variant
  Dim strGDBNewCoverFClassName As String
  Dim booGDBCoverHasFields As Boolean
  Dim lngGDBCoverFClassIndex As Long
  Dim lngGDBCoverQuadratIndex As Long
  Dim lngGDBCoverYearIndex As Long
  Dim lngGDBCoverTypeIndex As Long
  Dim lngGDBCoverOrigFIDIndex As Long
  Dim lngGDBCoverRotationIndex As Long

  Dim strYear As String
  Dim strQuadrat As String
  Dim strFClassName As String
  Dim strType As String

  Dim pSrcFCursor As IFeatureCursor
  Dim pSrcFeature As IFeature
  Dim pDestFCursor As IFeatureCursor
  Dim pDestFBuffer As IFeatureBuffer
  Dim pDestGDBFCursor As IFeatureCursor
  Dim pDestGDBFBuffer As IFeatureBuffer

  Dim pDestFClass As IFeatureClass
  Dim pDestGDBFClass As IFeatureClass
  Dim pPoint As IPoint
  Dim pPolygon As IPolygon
  Dim pClone As IClone
  Dim varIndexArray() As Variant
  Dim varGDBIndexArray() As Variant
  Dim lngFClassIndex As Long
  Dim lngQuadratIndex As Long
  Dim lngYearIndex As Long
  Dim lngTypeIndex As Long
  Dim lngOrigFIDIndex As Long
  Dim lngIsEmptyIndex As Long
  Dim lngRotationIndex As Long
  Dim lngGDBFClassIndex As Long
  Dim lngGDBQuadratIndex As Long
  Dim lngGDBYearIndex As Long
  Dim lngGDBTypeIndex As Long
  Dim lngGDBOrigFIDIndex As Long
  Dim lngGDBIsEmptyIndex As Long
  Dim lngGDBRotationIndex As Long

  Dim varCoverIndexArray() As Variant
  Dim varCoverGDBIndexArray() As Variant
  Dim varDensityIndexArray() As Variant
  Dim varDensityGDBIndexArray() As Variant

  Dim pNewCombinedDensityFClass As IFeatureClass
  Dim varCombinedDensityFieldIndexArray() As Variant
  Dim strNewCombinedDensityFClassName As String
  Dim booCombinedDensityHasFields As Boolean
  Dim lngCombinedDensityFClassIndex As Long
  Dim lngCombinedDensityQuadratIndex As Long
  Dim lngCombinedDensityYearIndex As Long
  Dim lngCombinedDensityTypeIndex As Long
  Dim lngCombinedDensityOrigFIDIndex As Long
  Dim lngCombinedDensityRotationIndex As Long

  Dim pNewCombinedCoverFClass As IFeatureClass
  Dim varCombinedCoverFieldIndexArray() As Variant
  Dim strNewCombinedCoverFClassName As String
  Dim booCombinedCoverHasFields As Boolean
  Dim lngCombinedCoverFClassIndex As Long
  Dim lngCombinedCoverQuadratIndex As Long
  Dim lngCombinedCoverYearIndex As Long
  Dim lngCombinedCoverTypeIndex As Long
  Dim lngCombinedCoverOrigFIDIndex As Long
  Dim lngCombinedCoverRotationIndex As Long

  Dim pCombinedDestFClass As IFeatureClass
  Dim varCombinedIndexArray() As Variant
  Dim lngCombinedFClassIndex As Long
  Dim lngCombinedQuadratIndex As Long
  Dim lngCombinedYearIndex As Long
  Dim lngCombinedTypeIndex As Long
  Dim lngCombinedOrigFIDIndex As Long
  Dim lngCombinedIsEmptyIndex As Long
  Dim lngCombinedRotationIndex As Long

  Dim pCombinedFCursor As IFeatureCursor
  Dim pCombinedFBuffer As IFeatureBuffer
  Dim pCombinedDensityFCursor As IFeatureCursor
  Dim pCombinedDensityFBuffer As IFeatureBuffer
  Dim pCombinedCoverFCursor As IFeatureCursor
  Dim pCombinedCoverFBuffer As IFeatureBuffer

  Dim strIsEmpty As String
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.000001

  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pTopoOp As ITopologicalOperator4

  Dim pTempDataset As IDataset
  Dim pTempFClass As IFeatureClass
  Dim strCoverType As String
  Dim strDensityType As String
  Dim strAltType As String
  Dim pAltDestFClass As IFeatureClass
  Dim varAltIndexArray() As Variant
  Dim lngAltFClassIndex As Long
  Dim lngAltQuadratIndex As Long
  Dim lngAltYearIndex As Long
  Dim lngAltIsEmptyIndex As Long
  Dim lngAltTypeIndex As Long
  Dim lngAltRotationIndex As Long

  Dim pAltDestGDBFClass As IFeatureClass
  Dim varAltGDBIndexArray() As Variant
  Dim lngAltGDBFClassIndex As Long
  Dim lngAltGDBQuadratIndex As Long
  Dim lngAltGDBYearIndex As Long
  Dim lngAltGDBTypeIndex As Long
  Dim lngAltGDBIsEmptyIndex As Long
  Dim lngAltGDBRotationIndex As Long

  Dim pAltCombinedDestFClass As IFeatureClass
  Dim varAltCombinedIndexArray() As Variant
  Dim lngAltCombinedFClassIndex As Long
  Dim lngAltCombinedQuadratIndex As Long
  Dim lngAltCombinedYearIndex As Long
  Dim lngAltCombinedTypeIndex As Long
  Dim lngAltCombinedIsEmptyIndex As Long
  Dim lngAltCombinedRotationIndex As Long

  Dim pAltCombinedFCursor As IFeatureCursor
  Dim pAltCombinedFBuffer As IFeatureBuffer

  Dim pAltDestFCursor As IFeatureCursor
  Dim pAltDestFBuffer As IFeatureBuffer
  Dim pAltDestGDBFCursor As IFeatureCursor
  Dim pAltDestGDBFBuffer As IFeatureBuffer

  Dim var_C_to_D_IndexArray() As Variant
  Dim var_D_to_C_IndexArray() As Variant

  Dim strSpecies As String
  Dim lngSpeciesIndex As Long
  Dim strHexSpecies As String
  Dim booShouldChange As Boolean
  Dim varPoints() As Variant
  Dim pTestPolygon As IPolygon
  Dim pTestPoint As IPoint
  Dim lngConvertIndex As Long
  Dim strCorD As String

  Dim booFoundPointInConversion As Boolean

  Dim pQuadrat As IPolygon
  Set pQuadrat = ReturnQuadratPolygon(pSpRef)
  Dim pNewPoly As IPolygon

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Set pNewFGDBWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strNewRoot & "\Combined_by_Site")
  Set pNewFeatFGDBWS = pNewFGDBWS

  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)

    varDatasets = ReturnFeatureClassesOrNothing(strFolder, booFoundShapefiles, booFoundPolys, booFoundPoints, _
        pRepPointFClass, pRepPolyFClass)

    booFoundPolys = True
    booFoundPoints = True

    Debug.Print CStr(lngIndex + 1) & " of " & CStr(pFolders.Count) & "] " & strFolder
    If booFoundShapefiles Then
      Debug.Print "--> Found Shapefiles = " & CStr(booFoundShapefiles) & " [n = " & CStr(UBound(varDatasets) + 1) & "]"

      strFolderName = aml_func_mod.ReturnFilename(strFolder)
      strNewFolder = strNewRoot & "\Shapefiles\" & strFolderName

      If Not aml_func_mod.ExistFileDir(strNewFolder) Then
        MyGeneralOperations.CreateNestedFoldersByPath strNewFolder
      End If
      Set pNewWS = pNewWSFact.OpenFromFile(strNewFolder, 0)
      Set pNewFeatWS = pNewWS

      If booFoundPoints Then
        Set pDataset = pRepPointFClass
        Call FillQuadratAndYearFromDatasetBrowsename(pDataset.BrowseName, strQuadrat, strYear, strCorD)
        strNewDensityFClassName = strQuadrat & "_Density"

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewDensityFClassName) Then
          Set pDataset = pNewFeatWS.OpenFeatureClass(strNewDensityFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varDensityFieldIndexArray
        Set pNewDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewWS, _
            varDensityFieldIndexArray, strNewDensityFClassName, booDensityHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewDensityFClass, lngDensityFClassIndex, lngDensityQuadratIndex, _
            lngDensityYearIndex, lngDensityTypeIndex, lngDensityOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewDensityFClass, strAbstract, strPurpose)
        DoEvents

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, strNewDensityFClassName) Then
          Set pDataset = pNewFeatFGDBWS.OpenFeatureClass(strNewDensityFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varGDBDensityFieldIndexArray
        Set pNewGDBDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewFGDBWS, _
            varGDBDensityFieldIndexArray, strNewDensityFClassName, booGDBDensityHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewGDBDensityFClass, lngGDBDensityFClassIndex, lngGDBDensityQuadratIndex, _
            lngGDBDensityYearIndex, lngGDBDensityTypeIndex, lngGDBDensityOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewGDBDensityFClass, strAbstract, strPurpose)
        DoEvents

        If pNewCombinedDensityFClass Is Nothing Then
          If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, "Density_All") Then
            Set pDataset = pNewFeatFGDBWS.OpenFeatureClass("Density_All")
            pDataset.DELETE
            Set pDataset = Nothing
          End If

          Erase varCombinedDensityFieldIndexArray
          Set pNewCombinedDensityFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPointFClass, pNewFGDBWS, _
              varCombinedDensityFieldIndexArray, "Density_All", booCombinedDensityHasFields, esriGeometryPolygon)

          Call CreateNewFields(pNewCombinedDensityFClass, lngCombinedDensityFClassIndex, lngCombinedDensityQuadratIndex, _
              lngCombinedDensityYearIndex, lngCombinedDensityTypeIndex, lngCombinedDensityOrigFIDIndex)
          Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCombinedDensityFClass, strAbstract, strPurpose)
          DoEvents

          Set pCombinedDensityFCursor = pNewCombinedDensityFClass.Insert(True)
          Set pCombinedDensityFBuffer = pNewCombinedDensityFClass.CreateFeatureBuffer
        End If

      End If

      If booFoundPolys Then
        Set pDataset = pRepPolyFClass
        Call FillQuadratAndYearFromDatasetBrowsename(pDataset.BrowseName, strQuadrat, strYear, strCorD)
        strNewCoverFClassName = strQuadrat & "_Cover"

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewWS, strNewCoverFClassName) Then
          Set pDataset = pNewFeatWS.OpenFeatureClass(strNewCoverFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varCoverFieldIndexArray
        Set pNewCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewWS, _
            varCoverFieldIndexArray, strNewCoverFClassName, booCoverHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewCoverFClass, lngCoverFClassIndex, lngCoverQuadratIndex, _
            lngCoverYearIndex, lngCoverTypeIndex, lngCoverOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCoverFClass, strAbstract, strPurpose)
        DoEvents

        If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, strNewCoverFClassName) Then
          Set pDataset = pNewFeatFGDBWS.OpenFeatureClass(strNewCoverFClassName)
          pDataset.DELETE
          Set pDataset = Nothing
        End If

        Erase varGDBCoverFieldIndexArray
        Set pNewGDBCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewFGDBWS, _
            varGDBCoverFieldIndexArray, strNewCoverFClassName, booGDBCoverHasFields, esriGeometryPolygon)

        Call CreateNewFields(pNewGDBCoverFClass, lngGDBCoverFClassIndex, lngGDBCoverQuadratIndex, _
            lngGDBCoverYearIndex, lngGDBCoverTypeIndex, lngGDBCoverOrigFIDIndex)
        Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewGDBCoverFClass, strAbstract, strPurpose)
        DoEvents

        If pNewCombinedCoverFClass Is Nothing Then
          If MyGeneralOperations.CheckIfFeatureClassExists(pNewFGDBWS, "Cover_All") Then
            Set pDataset = pNewFeatFGDBWS.OpenFeatureClass("Cover_All")
            pDataset.DELETE
            Set pDataset = Nothing
          End If

          Erase varCombinedCoverFieldIndexArray
          Set pNewCombinedCoverFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pRepPolyFClass, pNewFGDBWS, _
              varCombinedCoverFieldIndexArray, "Cover_All", booCombinedCoverHasFields, esriGeometryPolygon)

          Call CreateNewFields(pNewCombinedCoverFClass, lngCombinedCoverFClassIndex, lngCombinedCoverQuadratIndex, _
              lngCombinedCoverYearIndex, lngCombinedCoverTypeIndex, lngCombinedCoverOrigFIDIndex)
          Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pNewCombinedCoverFClass, strAbstract, strPurpose)
          DoEvents

          Set pCombinedCoverFCursor = pNewCombinedCoverFClass.Insert(True)
          Set pCombinedCoverFBuffer = pNewCombinedCoverFClass.CreateFeatureBuffer

        End If
      End If

      lngShapefileCount = lngShapefileCount + UBound(varDatasets) + 1

      For lngDatasetIndex = 0 To UBound(varDatasets)
        DoEvents

        Set pDataset = varDatasets(lngDatasetIndex)
        strFClassName = pDataset.BrowseName
        Call FillQuadratAndYearFromDatasetBrowsename(strFClassName, strQuadrat, strYear, strCorD)

        Set pFClass = pDataset

        lngSpeciesIndex = pFClass.FindField("Species")
        If lngSpeciesIndex = -1 Then
          DoEvents
        End If

        Debug.Print "  --> Adding Dataset '" & pDataset.BrowseName & "'"

        If strYear = 2022 And strQuadrat = "Q2" Then
          DoEvents
        End If

        If MyGeneralOperations.CheckCollectionForKey(pRotateColl, Replace(strQuadrat, "Q", "", , , vbTextCompare)) Then
          Set pCollByQuadrat = pRotateColl.Item(Replace(strQuadrat, "Q", "", , , vbTextCompare))
          varRotateElements = pCollByQuadrat.Item(strYear)
          strRotateBy = varRotateElements(3)
        Else
          strRotateBy = "0"
        End If

        If strFClassName = "Q67_2012_D" Then
          DoEvents
        End If

        If strCorD = "C" Then ' strSplit(UBound(strSplit)) = "C" Then
          strType = "Cover"
          Set pDestFClass = pNewCoverFClass
          varIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          lngFClassIndex = lngCoverFClassIndex
          lngQuadratIndex = lngCoverQuadratIndex
          lngYearIndex = lngCoverYearIndex
          lngTypeIndex = lngCoverTypeIndex
          lngIsEmptyIndex = pDestFClass.FindField("IsEmpty")
          lngRotationIndex = pDestFClass.FindField("Revise_Rtn")

          Set pDestGDBFClass = pNewGDBCoverFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBCoverFClassIndex
          lngGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngGDBYearIndex = lngGDBCoverYearIndex
          lngGDBTypeIndex = lngGDBCoverTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          lngGDBRotationIndex = pDestGDBFClass.FindField("Revise_Rtn")

          Set pCombinedDestFClass = pNewCombinedCoverFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngCombinedYearIndex = lngCombinedCoverYearIndex
          lngCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          lngCombinedRotationIndex = pCombinedDestFClass.FindField("Revise_Rtn")
          Set pCombinedFCursor = pCombinedCoverFCursor
          Set pCombinedFBuffer = pCombinedCoverFBuffer

          strAltType = "Density"
          Set pAltDestFClass = pNewDensityFClass
          ReDim varAltIndexArray(3, 4)

          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngDensityFClassIndex
          lngAltQuadratIndex = lngDensityQuadratIndex
          lngAltYearIndex = lngDensityYearIndex
          lngAltTypeIndex = lngDensityTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          lngAltRotationIndex = pAltDestFClass.FindField("Revise_Rtn")

          Set pAltDestGDBFClass = pNewGDBDensityFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBDensityFClassIndex
          lngAltGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngAltGDBYearIndex = lngGDBDensityYearIndex
          lngAltGDBTypeIndex = lngGDBDensityTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          lngAltGDBRotationIndex = pAltDestGDBFClass.FindField("Revise_Rtn")

          Set pAltCombinedDestFClass = pNewCombinedDensityFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngAltCombinedYearIndex = lngCombinedDensityYearIndex
          lngAltCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
          lngAltCombinedRotationIndex = pAltCombinedDestFClass.FindField("Revise_Rtn")
          Set pAltCombinedFCursor = pCombinedDensityFCursor
          Set pAltCombinedFBuffer = pCombinedDensityFBuffer

        Else
          strType = "Density"
          Set pDestFClass = pNewDensityFClass
          varIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestFClass)
          lngFClassIndex = lngDensityFClassIndex
          lngQuadratIndex = lngDensityQuadratIndex
          lngYearIndex = lngDensityYearIndex
          lngTypeIndex = lngDensityTypeIndex
          lngIsEmptyIndex = pDestFClass.FindField("IsEmpty")
          lngRotationIndex = pDestFClass.FindField("Revise_Rtn")

          Set pDestGDBFClass = pNewGDBDensityFClass
          varGDBIndexArray = ReturnArrayOfFieldLinks(pFClass, pDestGDBFClass)
          lngGDBFClassIndex = lngGDBDensityFClassIndex
          lngGDBQuadratIndex = lngGDBDensityQuadratIndex
          lngGDBYearIndex = lngGDBDensityYearIndex
          lngGDBTypeIndex = lngGDBDensityTypeIndex
          lngGDBIsEmptyIndex = pDestGDBFClass.FindField("IsEmpty")
          lngGDBRotationIndex = pDestGDBFClass.FindField("Revise_Rtn")

          Set pCombinedDestFClass = pNewCombinedDensityFClass
          varCombinedIndexArray = ReturnArrayOfFieldLinks(pFClass, pCombinedDestFClass)
          lngCombinedFClassIndex = lngCombinedDensityFClassIndex
          lngCombinedQuadratIndex = lngCombinedDensityQuadratIndex
          lngCombinedYearIndex = lngCombinedDensityYearIndex
          lngCombinedTypeIndex = lngCombinedDensityTypeIndex
          lngCombinedIsEmptyIndex = pCombinedDestFClass.FindField("IsEmpty")
          lngCombinedRotationIndex = pCombinedDestFClass.FindField("Revise_Rtn")
          Set pCombinedFCursor = pCombinedDensityFCursor
          Set pCombinedFBuffer = pCombinedDensityFBuffer

          strAltType = "Cover"
          Set pAltDestFClass = pNewCoverFClass
          varAltIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestFClass)
          lngAltFClassIndex = lngCoverFClassIndex
          lngAltQuadratIndex = lngCoverQuadratIndex
          lngAltYearIndex = lngCoverYearIndex
          lngAltTypeIndex = lngCoverTypeIndex
          lngAltIsEmptyIndex = pAltDestFClass.FindField("IsEmpty")
          lngAltRotationIndex = pAltDestFClass.FindField("Revise_Rtn")

          Set pAltDestGDBFClass = pNewGDBCoverFClass
          varAltGDBIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltDestGDBFClass)
          lngAltGDBFClassIndex = lngGDBCoverFClassIndex
          lngAltGDBQuadratIndex = lngGDBCoverQuadratIndex
          lngAltGDBYearIndex = lngGDBCoverYearIndex
          lngAltGDBTypeIndex = lngGDBCoverTypeIndex
          lngAltGDBIsEmptyIndex = pAltDestGDBFClass.FindField("IsEmpty")
          lngAltGDBRotationIndex = pAltDestGDBFClass.FindField("Revise_Rtn")

          Set pAltCombinedDestFClass = pNewCombinedCoverFClass
          varAltCombinedIndexArray = ReturnArrayOfFieldCrossLinks(pFClass, pAltCombinedDestFClass)
          lngAltCombinedFClassIndex = lngCombinedCoverFClassIndex
          lngAltCombinedQuadratIndex = lngCombinedCoverQuadratIndex
          lngAltCombinedYearIndex = lngCombinedCoverYearIndex
          lngAltCombinedTypeIndex = lngCombinedCoverTypeIndex
          lngAltCombinedIsEmptyIndex = pAltCombinedDestFClass.FindField("IsEmpty")
          lngAltCombinedRotationIndex = pAltCombinedDestFClass.FindField("Revise_Rtn")
          Set pAltCombinedFCursor = pCombinedCoverFCursor
          Set pAltCombinedFBuffer = pCombinedCoverFBuffer
        End If

        Set pDestFCursor = pDestFClass.Insert(True)
        Set pDestFBuffer = pDestFClass.CreateFeatureBuffer
        Set pDestGDBFCursor = pDestGDBFClass.Insert(True)
        Set pDestGDBFBuffer = pDestGDBFClass.CreateFeatureBuffer

        Set pAltDestFCursor = pAltDestFClass.Insert(True)
        Set pAltDestFBuffer = pAltDestFClass.CreateFeatureBuffer
        Set pAltDestGDBFCursor = pAltDestGDBFClass.Insert(True)
        Set pAltDestGDBFBuffer = pAltDestGDBFClass.CreateFeatureBuffer

        If pFClass.FindField("Cover") > -1 Or pFClass.FindField("Species") > -1 Then

          Set pSrcFCursor = pFClass.Search(Nothing, False)
          Set pSrcFeature = pSrcFCursor.NextFeature
          Do Until pSrcFeature Is Nothing
            strSpecies = pSrcFeature.Value(lngSpeciesIndex)

            strSpecies = Replace(strSpecies, vbCrLf, "", , , vbTextCompare)
            strSpecies = Replace(strSpecies, vbNewLine, "", , , vbTextCompare)
            strSpecies = Replace(strSpecies, vbTab, " ", , , vbTextCompare)
            strSpecies = Trim(strSpecies)
            Do Until InStr(1, strSpecies, "  ") = 0
              strSpecies = Replace(strSpecies, "  ", " ", , , vbTextCompare)
            Loop

            If strSpecies = "Menodora scabra" Then
              DoEvents
            End If
                strSpecies, strNoteOnChanges)

            If strSpecies = "Lycurus phleoides" Then
              DoEvents
            End If
            If Trim(strSpecies) = "" And strType = "Density" Then strSpecies = "No Point Species"
            If Trim(strSpecies) = "" And strType = "Cover" Then strSpecies = "No Polygon Species"

            strHexSpecies = HexifyName(strSpecies)
            If strType = "Density" Then
              If Not MyGeneralOperations.CheckCollectionForKey(pDensityShouldChangeColl, strHexSpecies) Then
                Debug.Print "Failed to find '" & strSpecies & "'..."

                lngFileNum = FreeFile(0)
                Open strMissingSpeciesPath For Append As lngFileNum
                Print #lngFileNum, """" & pDataset.BrowseName & """,""" & strSpecies & """"
                Close #lngFileNum

                If Not MyGeneralOperations.CheckCollectionForKey(pSpeciesSummaryColl, strSpecies) Then
                  Set pSubColl = New Collection
                  pSubColl.Add True, pDataset.BrowseName
                  ReDim strSubNames(0)
                  strSubNames(0) = pDataset.BrowseName
                  varSubArray = Array(strSubNames, pSubColl)
                  pSpeciesSummaryColl.Add varSubArray, strSpecies

                  If Not IsDimmed(strSpeciesArray) Then
                    ReDim strSpeciesArray(0)
                  Else
                    ReDim Preserve strSpeciesArray(UBound(strSpeciesArray) + 1)
                  End If
                  strSpeciesArray(UBound(strSpeciesArray)) = strSpecies
                Else
                  varSubArray = pSpeciesSummaryColl.Item(strSpecies)
                  strSubNames = varSubArray(0)
                  Set pSubColl = varSubArray(1)
                  If Not MyGeneralOperations.CheckCollectionForKey(pSubColl, pDataset.BrowseName) Then
                    ReDim Preserve strSubNames(UBound(strSubNames) + 1)
                    strSubNames(UBound(strSubNames)) = pDataset.BrowseName
                    pSubColl.Add True, pDataset.BrowseName
                    varSubArray = Array(strSubNames, pSubColl)
                    pSpeciesSummaryColl.Remove strSpecies
                    pSpeciesSummaryColl.Add varSubArray, strSpecies
                  End If
                End If

                booShouldChange = False
              Else
                booShouldChange = pDensityShouldChangeColl.Item(strHexSpecies)
              End If
            Else
              If Not MyGeneralOperations.CheckCollectionForKey(pCoverShouldChangeColl, strHexSpecies) Then
                Debug.Print "Failed to find '" & strSpecies & "'..."

                lngFileNum = FreeFile(0)
                Open strMissingSpeciesPath For Append As lngFileNum
                Print #lngFileNum, """" & pDataset.BrowseName & """,""" & strSpecies & """"
                Close #lngFileNum

                If Not MyGeneralOperations.CheckCollectionForKey(pSpeciesSummaryColl, strSpecies) Then
                  Set pSubColl = New Collection
                  pSubColl.Add True, pDataset.BrowseName
                  ReDim strSubNames(0)
                  strSubNames(0) = pDataset.BrowseName
                  varSubArray = Array(strSubNames, pSubColl)
                  pSpeciesSummaryColl.Add varSubArray, strSpecies

                  If Not IsDimmed(strSpeciesArray) Then
                    ReDim strSpeciesArray(0)
                  Else
                    ReDim Preserve strSpeciesArray(UBound(strSpeciesArray) + 1)
                  End If
                  strSpeciesArray(UBound(strSpeciesArray)) = strSpecies
                Else
                  varSubArray = pSpeciesSummaryColl.Item(strSpecies)
                  strSubNames = varSubArray(0)
                  Set pSubColl = varSubArray(1)
                  If Not MyGeneralOperations.CheckCollectionForKey(pSubColl, pDataset.BrowseName) Then
                    ReDim Preserve strSubNames(UBound(strSubNames) + 1)
                    strSubNames(UBound(strSubNames)) = pDataset.BrowseName
                    pSubColl.Add True, pDataset.BrowseName
                    varSubArray = Array(strSubNames, pSubColl)
                    pSpeciesSummaryColl.Remove strSpecies
                    pSpeciesSummaryColl.Add varSubArray, strSpecies
                  End If
                End If

                booShouldChange = False
              Else
                booShouldChange = pCoverShouldChangeColl.Item(strHexSpecies)
              End If
            End If

            If strType = "Density" Then
              Set pPoint = pSrcFeature.ShapeCopy
              If pPoint.IsEmpty Then
                Set pPolygon = New Polygon
              Else
                Set pPolygon = ReturnCircleClippedToQuadrat(pPoint, 0.001, 30, pQuadrat)
              End If
            Else
              Set pPolygon = pSrcFeature.ShapeCopy
            End If

            strIsEmpty = CBool(pPolygon.IsEmpty)
            Set pPolygon.SpatialReference = pSpRef

            Select Case strRotateBy
              Case "", "0"
                dblRotateVal = 0
              Case "CW 90"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(-90)   ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = -90
              Case "CCW 90"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(90)    ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = 90
              Case "180"
                Set pRotator = pPolygon
                pRotator.Rotate pMidPoint, MyGeometricOperations.DegToRad(180)   ' ASSUMING MATHEMATICAL ANGLES
                dblRotateVal = 180
              Case Else
                MsgBox "Unexpected Rotation! [" & strRotateBy & "]"
            End Select

            Set pTopoOp = pPolygon
            pTopoOp.IsKnownSimple = False
            pTopoOp.Simplify
            Set pClone = pPolygon

            Erase varPoints
            If booShouldChange Then
              If strType = "Cover" Then
                If pPolygon.IsEmpty Then
                  ReDim varPoints(0)
                  Set pTestPolygon = New Polygon
                  Set pTestPolygon.SpatialReference = pSpRef
                  Set varPoints(0) = pTestPolygon
                Else
                  varPoints = Margaret_Functions.FillPolygonWithPointArray(pPolygon, 0.005)
                  booFoundPointInConversion = False
                  For lngConvertIndex = 0 To UBound(varPoints)
                    Set pTestPoint = varPoints(lngConvertIndex)
                    Set pTestPoint.SpatialReference = pSpRef
                    Set pTestPolygon = ReturnCircleClippedToQuadrat(pTestPoint, 0.001, 30, pQuadrat, pPolygon)
                    Set pTestPolygon.SpatialReference = pSpRef

                    Set varPoints(lngConvertIndex) = pTestPolygon
                    If Not pTestPolygon.IsEmpty Then
                      booFoundPointInConversion = True
                    End If
                  Next lngConvertIndex
                  If Not booFoundPointInConversion Then
                    DoEvents
                  Else
                  End If

                End If
              Else  ' IF STARTING AS DENSITY AND CONVERTING TO COVER; DECIDE IF WE WANT TO MAKE THIS A BIGGER POLYGON
                ReDim varPoints(0)
                Set varPoints(0) = pClone.Clone
              End If
            End If

            If booShouldChange Then
              For lngConvertIndex = 0 To UBound(varPoints)
                Set pClone = varPoints(lngConvertIndex)
                Set pNewPoly = pClone.Clone
                If Not pNewPoly.IsEmpty Then

                  Set pAltDestFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltIndexArray, 2)
                    pAltDestFBuffer.Value(varAltIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltDestFBuffer.Value(lngAltFClassIndex) = strFClassName
                  pAltDestFBuffer.Value(lngAltQuadratIndex) = strQuadrat
                  pAltDestFBuffer.Value(lngAltYearIndex) = strYear
                  pAltDestFBuffer.Value(lngAltTypeIndex) = strAltType
                  pAltDestFBuffer.Value(lngAltIsEmptyIndex) = strIsEmpty
                  pAltDestFBuffer.Value(lngAltRotationIndex) = dblRotateVal
                  pAltDestFCursor.InsertFeature pAltDestFBuffer

                  Set pAltDestGDBFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltGDBIndexArray, 2)
                    pAltDestGDBFBuffer.Value(varAltGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltGDBIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltDestGDBFBuffer.Value(lngAltGDBFClassIndex) = strFClassName
                  pAltDestGDBFBuffer.Value(lngAltGDBQuadratIndex) = strQuadrat
                  pAltDestGDBFBuffer.Value(lngAltGDBYearIndex) = strYear
                  pAltDestGDBFBuffer.Value(lngAltGDBTypeIndex) = strAltType
                  pAltDestGDBFBuffer.Value(lngAltGDBIsEmptyIndex) = strIsEmpty
                  pAltDestGDBFBuffer.Value(lngAltGDBRotationIndex) = dblRotateVal
                  pAltDestGDBFCursor.InsertFeature pAltDestGDBFBuffer

                  Set pAltCombinedFBuffer.Shape = pClone.Clone
                  For lngIndex2 = 0 To UBound(varAltCombinedIndexArray, 2)
                    pAltCombinedFBuffer.Value(varAltCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varAltCombinedIndexArray(1, lngIndex2))
                  Next lngIndex2
                  pAltCombinedFBuffer.Value(lngAltCombinedFClassIndex) = strFClassName
                  pAltCombinedFBuffer.Value(lngAltCombinedQuadratIndex) = strQuadrat
                  pAltCombinedFBuffer.Value(lngAltCombinedYearIndex) = strYear
                  pAltCombinedFBuffer.Value(lngAltCombinedTypeIndex) = strAltType
                  pAltCombinedFBuffer.Value(lngAltCombinedIsEmptyIndex) = strIsEmpty
                  pAltCombinedFBuffer.Value(lngAltCombinedRotationIndex) = dblRotateVal
                  pAltCombinedFCursor.InsertFeature pAltCombinedFBuffer
                End If
              Next lngConvertIndex

            Else
              Set pDestFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                If IsNull(pSrcFeature.Value(varIndexArray(1, lngIndex2))) Then
                  If pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Type = esriFieldTypeString Then
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = ""
                  Else
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = 0
                  End If
                Else
                  If varIndexArray(3, lngIndex2) > -1 Then
                    pDestFBuffer.Value(varIndexArray(3, lngIndex2)) = pSrcFeature.Value(varIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pDestFBuffer.Value(lngFClassIndex) = strFClassName
              pDestFBuffer.Value(lngQuadratIndex) = strQuadrat
              pDestFBuffer.Value(lngYearIndex) = strYear
              pDestFBuffer.Value(lngTypeIndex) = strType
              pDestFBuffer.Value(lngIsEmptyIndex) = strIsEmpty
              pDestFBuffer.Value(lngRotationIndex) = dblRotateVal
              pDestFCursor.InsertFeature pDestFBuffer

              Set pDestGDBFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varGDBIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"
                If varGDBIndexArray(3, lngIndex2) <> -1 Then
                  If pDestGDBFBuffer.Fields.Field(varGDBIndexArray(3, lngIndex2)).Editable Then
                    pDestGDBFBuffer.Value(varGDBIndexArray(3, lngIndex2)) = pSrcFeature.Value(varGDBIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pDestGDBFBuffer.Value(lngGDBFClassIndex) = strFClassName
              pDestGDBFBuffer.Value(lngGDBQuadratIndex) = strQuadrat
              pDestGDBFBuffer.Value(lngGDBYearIndex) = strYear
              pDestGDBFBuffer.Value(lngGDBTypeIndex) = strType
              pDestGDBFBuffer.Value(lngGDBIsEmptyIndex) = strIsEmpty
              pDestGDBFBuffer.Value(lngGDBRotationIndex) = dblRotateVal
              pDestGDBFCursor.InsertFeature pDestGDBFBuffer

              Set pCombinedFBuffer.Shape = pClone.Clone
              For lngIndex2 = 0 To UBound(varCombinedIndexArray, 2)
                  " from Source [" & CStr(varIndexArray(0, lngIndex2)) & _
                  ", Index " & CStr(varIndexArray(1, lngIndex2)) & _
                  ", Fieldname = '" & pSrcFeature.Fields.Field(varIndexArray(1, lngIndex2)).Name & _
                  "'] to Destination [" & _
                  CStr(varIndexArray(2, lngIndex2)) & ", Index " & CStr(varIndexArray(3, lngIndex2)) & _
                  ", Fieldname = '" & pDestFBuffer.Fields.Field(varIndexArray(3, lngIndex2)).Name & _
                  "']"

                If varCombinedIndexArray(3, lngIndex2) <> -1 Then
                  If pCombinedFBuffer.Fields.Field(varCombinedIndexArray(3, lngIndex2)).Editable Then
                    pCombinedFBuffer.Value(varCombinedIndexArray(3, lngIndex2)) = pSrcFeature.Value(varCombinedIndexArray(1, lngIndex2))
                  End If
                End If
              Next lngIndex2
              pCombinedFBuffer.Value(lngCombinedFClassIndex) = strFClassName
              pCombinedFBuffer.Value(lngCombinedQuadratIndex) = strQuadrat
              pCombinedFBuffer.Value(lngCombinedYearIndex) = strYear
              pCombinedFBuffer.Value(lngCombinedTypeIndex) = strType
              pCombinedFBuffer.Value(lngCombinedIsEmptyIndex) = strIsEmpty
              pCombinedFBuffer.Value(lngCombinedRotationIndex) = dblRotateVal
              pCombinedFCursor.InsertFeature pCombinedFBuffer
            End If

            Set pSrcFeature = pSrcFCursor.NextFeature
          Loop

          pDestFCursor.Flush
          pDestGDBFCursor.Flush
          pCombinedFCursor.Flush

        End If
      Next lngDatasetIndex

      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCoverFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBCoverFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewDensityFClass, "Plot")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "SPCODE")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "FClassName")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Seedling")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Species")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Quadrat")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Year")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Type")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Orig_FID")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Verb_Spcs")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Site")
      Call MyGeneralOperations.CreateFieldAttributeIndex(pNewGDBDensityFClass, "Plot")

    End If

  Next lngIndex

  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedDensityFClass, "Plot")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Year")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pNewCombinedCoverFClass, "Plot")

  If IsDimmed(strSpeciesArray) Then
    QuickSort.StringsAscending strSpeciesArray, 0, UBound(strSpeciesArray)
    lngFileNum = FreeFile(0)
    Open strMissingSummaryPath For Append As lngFileNum
    For lngIndex = 0 To UBound(strSpeciesArray)
      strSpeciesLine = ""
      strSpecies = strSpeciesArray(lngIndex)
      varSubArray = pSpeciesSummaryColl.Item(strSpecies)
      strSubNames = varSubArray(0)
      If IsDimmed(strSubNames) Then
        For lngIndex2 = 0 To UBound(strSubNames)
          strSpeciesLine = strSpeciesLine & strSubNames(lngIndex2) & IIf(lngIndex2 = UBound(strSubNames), "", ", ")
        Next lngIndex2
        Print #lngFileNum, """" & strSpecies & """,""" & strSpeciesLine & """"
      End If
    Next lngIndex
    Close lngFileNum
  End If

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing
  Set pCoverToDensity = Nothing
  Set pDensityToCover = Nothing
  Set pCoverShouldChangeColl = Nothing
  Set pDensityShouldChangeColl = Nothing
  Set pRotateColl = Nothing
  Set pRotator = Nothing
  Set pCollByQuadrat = Nothing
  Erase varRotateElements
  Set pMidPoint = Nothing
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Erase strSpeciesArray
  Set pSpeciesSummaryColl = Nothing
  Set pSubColl = Nothing
  Erase strSubNames
  Erase varSubArray
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pCheckCollection = Nothing
  Set pFClass = Nothing
  Set pRepPointFClass = Nothing
  Set pRepPolyFClass = Nothing
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pField = Nothing
  Set pNewFields = Nothing
  Set pNewDensityFClass = Nothing
  Erase varDensityFieldIndexArray
  Set pNewGDBDensityFClass = Nothing
  Erase varGDBDensityFieldIndexArray
  Set pNewCoverFClass = Nothing
  Erase varCoverFieldIndexArray
  Set pNewGDBCoverFClass = Nothing
  Erase varGDBCoverFieldIndexArray
  Set pSrcFCursor = Nothing
  Set pSrcFeature = Nothing
  Set pDestFCursor = Nothing
  Set pDestFBuffer = Nothing
  Set pDestGDBFCursor = Nothing
  Set pDestGDBFBuffer = Nothing
  Set pDestFClass = Nothing
  Set pDestGDBFClass = Nothing
  Set pPoint = Nothing
  Set pPolygon = Nothing
  Set pClone = Nothing
  Erase varIndexArray
  Erase varGDBIndexArray
  Erase varCoverIndexArray
  Erase varCoverGDBIndexArray
  Erase varDensityIndexArray
  Erase varDensityGDBIndexArray
  Set pNewCombinedDensityFClass = Nothing
  Erase varCombinedDensityFieldIndexArray
  Set pNewCombinedCoverFClass = Nothing
  Erase varCombinedCoverFieldIndexArray
  Set pCombinedDestFClass = Nothing
  Erase varCombinedIndexArray
  Set pCombinedFCursor = Nothing
  Set pCombinedFBuffer = Nothing
  Set pCombinedDensityFCursor = Nothing
  Set pCombinedDensityFBuffer = Nothing
  Set pCombinedCoverFCursor = Nothing
  Set pCombinedCoverFBuffer = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pTopoOp = Nothing
  Set pTempDataset = Nothing
  Set pTempFClass = Nothing
  Set pAltDestFClass = Nothing
  Erase varAltIndexArray
  Set pAltDestGDBFClass = Nothing
  Erase varAltGDBIndexArray
  Set pAltCombinedDestFClass = Nothing
  Erase varAltCombinedIndexArray
  Set pAltCombinedFCursor = Nothing
  Set pAltCombinedFBuffer = Nothing
  Set pAltDestFCursor = Nothing
  Set pAltDestFBuffer = Nothing
  Set pAltDestGDBFCursor = Nothing
  Set pAltDestGDBFBuffer = Nothing
  Erase var_C_to_D_IndexArray
  Erase var_D_to_C_IndexArray
  Erase varPoints
  Set pTestPolygon = Nothing
  Set pTestPoint = Nothing
  Set pQuadrat = Nothing
  Set pNewPoly = Nothing

End Sub

Public Sub FillQuadratAndYearFromDatasetBrowsename(strBrowseName As String, strQuadratToFill As String, strYearToFill As String, _
    strCorDToFill As String)

  Dim strSplit() As String
  Dim lngIndex As Long
  strQuadratToFill = ""
  strYearToFill = ""
  strCorDToFill = ""

  Dim strFilename As String
  strFilename = aml_func_mod.ClipExtension2(strBrowseName)
  strSplit = Split(strFilename, "_")
  For lngIndex = 0 To UBound(strSplit)
    If Left(strSplit(lngIndex), 1) = "Q" Then
      strQuadratToFill = strSplit(lngIndex)
      strYearToFill = strSplit(lngIndex + 1)
      Exit For
    End If
  Next lngIndex
  If StrComp(strSplit(UBound(strSplit)), "C", vbTextCompare) = 0 Then
    strCorDToFill = "C"
  ElseIf StrComp(strSplit(UBound(strSplit)), "D", vbTextCompare) = 0 Then
    strCorDToFill = "D"
  End If

  If strQuadratToFill = "" Or strYearToFill = "" Or strCorDToFill = "" Then
           "strYearToFill = '" & strYearToFill & "'" & vbCrLf & _
           "strYearToFill = '" & strYearToFill & "'", , "Checking '" & strBrowseName & "'"
  End If

End Sub

Public Sub AddEmptyFeaturesAndFeatureClassesToCleaned_SA()

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Dim strRecreatedFolder As String

  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, , , , strRecreatedFolder, strContainingFolder)

  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
  Set pNewFeatWS = pNewWS
  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pNewFGDBWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Combined_by_Site.gdb", 0)
  Set pNewFeatFGDBWS = pNewFGDBWS

  Dim strEmptyFeatureReport As String
  Dim strEmptyKey As String
  Dim pDoneEmptyFeaturesColl As Collection
  Dim strEmptyYear As String
  Dim strQuad As String
  Dim pInsertCursor As IFeatureCursor
  Dim pInsertBuffer As IFeatureBuffer
  Dim pEmptyPolygon As IPolygon
  Dim varItems() As Variant
  Dim strSite As String
  Dim strPlot As String
  Dim lngStartYear As Long
  Dim lngEndYear As Long
  Dim booSurveyedThisYear As Boolean
  Dim lngEmptyYearIndex As Long
  Dim pQueryFilt As IQueryFilter
  Dim strShapefilePrefix As String
  Dim strShapefileSuffix As String
  Dim strGDBPrefix As String
  Dim strGDBSuffix As String
  Dim pYearsSiteSurveyed As Collection

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewWS, strShapefilePrefix, strShapefileSuffix)
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewFGDBWS, strGDBPrefix, strGDBSuffix)

  lngStartYear = 1935
  lngEndYear = 2028
  Dim pSitesSurveyedByYearColl As Collection
  Set pSitesSurveyedByYearColl = More_Margaret_Functions.ReturnCollectionOfYearsSurveyedByQuadrat(lngStartYear, lngEndYear)
  Set pQueryFilt = New QueryFilter
  Set pDoneEmptyFeaturesColl = New Collection
  strEmptyFeatureReport = """Quadrat""" & vbTab & """Site""" & vbTab & """Plot""" & vbTab & _
      """Year""" & vbTab & """Type""" & vbCrLf

  Dim pQuadData As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim strQuadratNames() As String
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev_SA(strQuadratNames, pPlotToQuadratConversion, _
      pQuadratToPlotConversion, varSites, varSitesSpecific, booRestrictToSite, strSiteToRestrict)

  Dim lngIndex As Long
  Dim strQuadrat As String
  Dim pNewCoverFClass As IFeatureClass
  Dim pNewDensityFClass As IFeatureClass
  Dim pSpRef As ISpatialReference
  Dim strFClassName As String
  Dim pNewGDBCoverFClass As IFeatureClass
  Dim pNewGDBDensityFClass As IFeatureClass
  Dim pNewGDBCoverAllFClass As IFeatureClass
  Dim pNewGDBDensityAllFClass As IFeatureClass
  Dim pGeoDataset As IGeoDataset

  Set pNewGDBCoverAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Cover_All")
  Set pNewGDBDensityAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Density_All")
  Set pGeoDataset = pNewGDBCoverAllFClass
  Set pSpRef = pGeoDataset.SpatialReference

  pSBar.ShowProgressBar "Adding Empty Features...", 0, UBound(strQuadratNames) + 1, 1, True
  pProg.position = 0
  Dim lngIndex2 As Long
  Dim strNewFClassName As String

  For lngIndex = 0 To UBound(strQuadratNames)
    DoEvents
    pProg.Step
    strQuadrat = "Q" & strQuadratNames(lngIndex)
    strQuad = Replace(strQuadrat, "Q", "", , , vbTextCompare)

    If strQuad <> "496" Then
      varItems = pQuadData.Item(strQuad)
      strSite = Trim(varItems(1))
      If strSite = "" Then
        strSite = Trim(varItems(0))
      End If
      strPlot = Trim(varItems(2))

      strNewFClassName = ReplaceBadChars(strSite, True, True, True, True)
      Do Until InStr(1, strNewFClassName, "__", vbTextCompare) = 0
        strNewFClassName = Replace(strNewFClassName, "__", "_")
      Loop
      Debug.Print "Reviewing Quadrat " & strQuadrat & " [from " & strNewFClassName & "]"

      Set pYearsSiteSurveyed = pSitesSurveyedByYearColl.Item(strQuadrat)

      Set pNewWSFact = New ShapefileWorkspaceFactory
      Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
      Set pNewFeatWS = pNewWS
      Set pNewCoverFClass = pNewFeatWS.OpenFeatureClass(strNewFClassName & "_Cover")
      Set pNewDensityFClass = pNewFeatWS.OpenFeatureClass(strNewFClassName & "_Density")

      Set pNewGDBCoverFClass = pNewFeatFGDBWS.OpenFeatureClass(strNewFClassName & "_Cover")
      Set pNewGDBDensityFClass = pNewFeatFGDBWS.OpenFeatureClass(strNewFClassName & "_Density")

      For lngEmptyYearIndex = lngStartYear To lngEndYear
        strEmptyYear = Format(lngEmptyYearIndex, "0")
        strFClassName = strQuadrat & "_" & strEmptyYear & "_"
        booSurveyedThisYear = pYearsSiteSurveyed.Item(strEmptyYear)

        If strEmptyYear = "2004" And strQuadrat = "Q84" Then
          DoEvents
        End If

        If booSurveyedThisYear Then
          If pNewCoverFClass.FindField("Year") = -1 And pNewCoverFClass.FindField("z_Year") > 0 Then
            pQueryFilt.WhereClause = strShapefilePrefix & "Quadrat" & strShapefileSuffix & " = '" & strQuadrat & "' AND " & _
                strShapefilePrefix & "z_Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"
          Else
            pQueryFilt.WhereClause = strShapefilePrefix & "Quadrat" & strShapefileSuffix & " = '" & strQuadrat & "' AND " & _
                strShapefilePrefix & "Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"
          End If

          If pNewCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewCoverFClass.Insert(True)
            Set pInsertBuffer = pNewCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("z_Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewDensityFClass.Insert(True)
            Set pInsertBuffer = pNewDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("z_Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
               strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
              strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

        End If
      Next lngEmptyYearIndex
    End If
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  strEmptyFeatureReport = Replace(strEmptyFeatureReport, vbTab, ",", , , vbTextCompare)

  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainingFolder & "\Empty_Features_Added_to_Cleaned_dataset.csv")
  MyGeneralOperations.WriteTextFile strExportPath, strEmptyFeatureReport, False, False

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFolders = Nothing
  Set pDoneEmptyFeaturesColl = Nothing
  Set pInsertCursor = Nothing
  Set pInsertBuffer = Nothing
  Set pEmptyPolygon = Nothing
  Erase varItems
  Set pQueryFilt = Nothing
  Set pYearsSiteSurveyed = Nothing
  Set pSitesSurveyedByYearColl = Nothing
  Set pQuadData = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Erase strQuadratNames
  Set pNewCoverFClass = Nothing
  Set pNewDensityFClass = Nothing
  Set pSpRef = Nothing
  Set pNewGDBCoverFClass = Nothing
  Set pNewGDBDensityFClass = Nothing
  Set pNewGDBCoverAllFClass = Nothing
  Set pNewGDBDensityAllFClass = Nothing
  Set pGeoDataset = Nothing

End Sub

Public Sub AddEmptyFeaturesAndFeatureClasses_SA(Optional booDoRecreated As Boolean = False)

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim strNewFolder As String
  Dim pNewWS As IWorkspace
  Dim pNewFeatWS As IFeatureWorkspace
  Dim pNewFGDBWS As IWorkspace
  Dim pNewFeatFGDBWS As IFeatureWorkspace
  Dim pNewWSFact As IWorkspaceFactory

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pApp As IApplication
  Set pApp = Application
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Dim strRecreatedFolder As String

  Dim strNewRoot As String
  Dim strExportPath As String

  Call DeclareWorkspaces(strRoot, strNewRoot, , , strRecreatedFolder, strContainingFolder)

  If booDoRecreated Then
    Set pNewWSFact = New ShapefileWorkspaceFactory
    Set pNewWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Shapefiles", 0)
    Set pNewFeatWS = pNewWS
    Set pNewWSFact = New FileGDBWorkspaceFactory
    Set pNewFGDBWS = pNewWSFact.OpenFromFile(strRecreatedFolder & "\Combined_by_Site.gdb", 0)
    Set pNewFeatFGDBWS = pNewFGDBWS
  Else
    Set pNewWSFact = New ShapefileWorkspaceFactory
    Set pNewWS = pNewWSFact.OpenFromFile(strNewRoot & "\Shapefiles", 0)
    Set pNewFeatWS = pNewWS
    Set pNewWSFact = New FileGDBWorkspaceFactory
    Set pNewFGDBWS = pNewWSFact.OpenFromFile(strNewRoot & "\Combined_by_Site.gdb", 0)
    Set pNewFeatFGDBWS = pNewFGDBWS
  End If

  Dim strEmptyFeatureReport As String
  Dim strEmptyKey As String
  Dim pDoneEmptyFeaturesColl As Collection
  Dim strEmptyYear As String
  Dim strQuad As String
  Dim pInsertCursor As IFeatureCursor
  Dim pInsertBuffer As IFeatureBuffer
  Dim pEmptyPolygon As IPolygon
  Dim varItems() As Variant
  Dim strSite As String
  Dim strPlot As String
  Dim lngStartYear As Long
  Dim lngEndYear As Long
  Dim booSurveyedThisYear As Boolean
  Dim lngEmptyYearIndex As Long
  Dim pQueryFilt As IQueryFilter
  Dim strShapefilePrefix As String
  Dim strShapefileSuffix As String
  Dim strGDBPrefix As String
  Dim strGDBSuffix As String
  Dim pYearsSiteSurveyed As Collection

  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewWS, strShapefilePrefix, strShapefileSuffix)
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pNewFGDBWS, strGDBPrefix, strGDBSuffix)

  lngStartYear = 1935
  lngEndYear = 2028
  Dim pSitesSurveyedByYearColl As Collection
  Set pSitesSurveyedByYearColl = More_Margaret_Functions.ReturnCollectionOfYearsSurveyedByQuadrat(lngStartYear, lngEndYear)
  Set pQueryFilt = New QueryFilter
  Set pDoneEmptyFeaturesColl = New Collection
  strEmptyFeatureReport = """Quadrat""" & vbTab & """Site""" & vbTab & """Plot""" & vbTab & _
      """Year""" & vbTab & """Type""" & vbCrLf

  Dim pQuadData As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim strQuadratNames() As String
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev_SA(strQuadratNames, pPlotToQuadratConversion, _
      pQuadratToPlotConversion, varSites, varSitesSpecific, booRestrictToSite, strSiteToRestrict)

  Dim lngIndex As Long
  Dim strQuadrat As String
  Dim pNewCoverFClass As IFeatureClass
  Dim pNewDensityFClass As IFeatureClass
  Dim pSpRef As ISpatialReference
  Dim strFClassName As String
  Dim pNewGDBCoverFClass As IFeatureClass
  Dim pNewGDBDensityFClass As IFeatureClass
  Dim pNewGDBCoverAllFClass As IFeatureClass
  Dim pNewGDBDensityAllFClass As IFeatureClass
  Dim pGeoDataset As IGeoDataset

  Set pNewGDBCoverAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Cover_All")
  Set pNewGDBDensityAllFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "Density_All")
  Set pGeoDataset = pNewGDBCoverAllFClass
  Set pSpRef = pGeoDataset.SpatialReference

  pSBar.ShowProgressBar "Adding Empty Features...", 0, UBound(strQuadratNames) + 1, 1, True
  pProg.position = 0

  For lngIndex = 0 To UBound(strQuadratNames)
    DoEvents
    pProg.Step
    strQuadrat = "Q" & strQuadratNames(lngIndex)
    strQuad = Replace(strQuadrat, "Q", "", , , vbTextCompare)

    If strQuad <> "496" Then
      varItems = pQuadData.Item(strQuad)
      strSite = Trim(varItems(1))
      If strSite = "" Then
        strSite = Trim(varItems(0))
      End If
      strPlot = Trim(varItems(2))
      Debug.Print "Reviewing Quadrat " & strQuadrat & "..."

      Set pYearsSiteSurveyed = pSitesSurveyedByYearColl.Item(strQuadrat)

      Set pNewWSFact = New ShapefileWorkspaceFactory
      Set pNewWS = pNewWSFact.OpenFromFile(strNewRoot & "\Shapefiles\" & Replace(strSite, " ", "_") & "_" & strQuadrat, 0)
      Set pNewFeatWS = pNewWS
      Set pNewCoverFClass = pNewFeatWS.OpenFeatureClass(strQuadrat & "_Cover")
      Set pNewDensityFClass = pNewFeatWS.OpenFeatureClass(strQuadrat & "_Density")

      Set pNewGDBCoverFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "_Cover")
      Set pNewGDBDensityFClass = pNewFeatFGDBWS.OpenFeatureClass(strQuadrat & "_Density")

      For lngEmptyYearIndex = lngStartYear To lngEndYear
        strEmptyYear = Format(lngEmptyYearIndex, "0")
        strFClassName = strQuadrat & "_" & strEmptyYear & "_"
        booSurveyedThisYear = pYearsSiteSurveyed.Item(strEmptyYear)
        If booSurveyedThisYear Then
          pQueryFilt.WhereClause = strShapefilePrefix & "Year" & strShapefileSuffix & " = '" & strEmptyYear & "'"

          If pNewCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewCoverFClass.Insert(True)
            Set pInsertBuffer = pNewCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewDensityFClass.Insert(True)
            Set pInsertBuffer = pNewDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = 0
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = ""
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = "-999"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          pQueryFilt.WhereClause = strGDBPrefix & "Quadrat" & strGDBSuffix & " = '" & strQuadrat & "' AND " & _
              strGDBPrefix & "Year" & strGDBSuffix & " = '" & strEmptyYear & "'"

          If pNewGDBCoverAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBCoverAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBCoverAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "C"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Cover"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Cover Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Cover""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

          If pNewGDBDensityAllFClass.FeatureCount(pQueryFilt) = 0 Then
            Set pInsertCursor = pNewGDBDensityAllFClass.Insert(True)
            Set pInsertBuffer = pNewGDBDensityAllFClass.CreateFeatureBuffer
            Set pEmptyPolygon = New Polygon
            Set pEmptyPolygon.SpatialReference = pSpRef
            Set pInsertBuffer.Shape = pEmptyPolygon
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("SPCODE")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("FClassName")) = strFClassName & "D"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Seedling")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Species")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Quadrat")) = strQuadrat
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Year")) = strEmptyYear
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Type")) = "Density"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Orig_FID")) = Null
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Verb_Spcs")) = "No Density Species Observed"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Site")) = strSite
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Plot")) = strPlot
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("IsEmpty")) = "True"
            pInsertBuffer.Value(pInsertBuffer.Fields.FindField("Revise_Rtn")) = 0

            pInsertCursor.InsertFeature pInsertBuffer
            pInsertCursor.Flush

            strEmptyKey = """" & strQuadrat & """" & vbTab & """" & strSite & """" & vbTab & """" & strPlot & """" & _
              vbTab & """" & strEmptyYear & """" & vbTab & """Density""" & vbCrLf
            If Not MyGeneralOperations.CheckCollectionForKey(pDoneEmptyFeaturesColl, strEmptyKey) Then
              pDoneEmptyFeaturesColl.Add True, strEmptyKey
              strEmptyFeatureReport = strEmptyFeatureReport & strEmptyKey
            End If
          End If

        End If
      Next lngEmptyYearIndex
    End If
  Next lngIndex

  pProg.position = 0
  pSBar.HideProgressBar

  strEmptyFeatureReport = Replace(strEmptyFeatureReport, vbTab, ",", , , vbTextCompare)

  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainingFolder & "\Empty_Features_Added.csv")
  MyGeneralOperations.WriteTextFile strExportPath, strEmptyFeatureReport, False, False

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pNewWS = Nothing
  Set pNewFeatWS = Nothing
  Set pNewFGDBWS = Nothing
  Set pNewFeatFGDBWS = Nothing
  Set pNewWSFact = Nothing
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pFolders = Nothing
  Set pDoneEmptyFeaturesColl = Nothing
  Set pInsertCursor = Nothing
  Set pInsertBuffer = Nothing
  Set pEmptyPolygon = Nothing
  Erase varItems
  Set pQueryFilt = Nothing
  Set pYearsSiteSurveyed = Nothing
  Set pSitesSurveyedByYearColl = Nothing
  Set pQuadData = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Erase strQuadratNames
  Set pNewCoverFClass = Nothing
  Set pNewDensityFClass = Nothing
  Set pSpRef = Nothing
  Set pNewGDBCoverFClass = Nothing
  Set pNewGDBDensityFClass = Nothing
  Set pNewGDBCoverAllFClass = Nothing
  Set pNewGDBDensityAllFClass = Nothing
  Set pGeoDataset = Nothing

End Sub

Public Function FillRotateColl_SA() As Collection

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Debug.Print "  --> Extracting Rotation Info..."
  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim varSites() As Variant
  Dim varSitesSpecific() As Variant

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument

  Dim strNewSource As String
  strNewSource = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Summary_Data_from_JSJ\Rotation.xlsx"

  Dim strFolder As String
  Dim lngIndex As Long

  Set pQuadratColl = FillQuadratNameColl_Rev_SA(strQuadratNames, pPlotToQuadratConversion, pQuadratToPlotConversion, _
       varSites, varSitesSpecific, booRestrictToSite, strSiteToRestrict)

  Dim pWSFact As IWorkspaceFactory
  Dim pWS As IFeatureWorkspace
  Set pWSFact = New ExcelWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strNewSource, 0)

  Dim pTable As ITable
  Set pTable = pWS.OpenTable("For_ArcGIS$")

  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim pReturn As New Collection
  Dim lngSiteIndex As Long
  Dim lngPlotIndex As Long
  Dim lngYearIndex As Long
  Dim lngTurnIndex As Long
  Dim lngNotesIndex As Long
  Dim lngExtraNotesIndex As Long

  lngSiteIndex = pTable.FindField("Site")
  lngPlotIndex = pTable.FindField("Quadrat")
  lngYearIndex = pTable.FindField("Year")
  lngTurnIndex = pTable.FindField("Turn_quadrat")
  lngNotesIndex = pTable.FindField("Notes")
  lngExtraNotesIndex = pTable.FindField("Extra_Notes")

  Dim strSite As String
  Dim strPlot As String
  Dim strQuadrat As String
  Dim strYear As String
  Dim strTurn As String
  Dim strNotes As String
  Dim strExtra As String
  Dim varElement() As Variant
  Dim varVal As Variant
  Dim pQuadratByYearColl As Collection

  Set pCursor = pTable.Search(Nothing, False)
  Set pRow = pCursor.NextRow
  Do Until pRow Is Nothing
    varVal = pRow.Value(lngSiteIndex)
    If IsNull(varVal) Then
      strSite = ""
    Else
      strSite = Trim(CStr(varVal))
    End If

    varVal = pRow.Value(lngPlotIndex)
    If IsNull(varVal) Then
      strPlot = ""
    Else
      strPlot = Trim(CStr(varVal))
    End If

    If MyGeneralOperations.CheckCollectionForKey(pPlotToQuadratConversion, strPlot) Then
    strQuadrat = pPlotToQuadratConversion.Item(strPlot)

      varVal = pRow.Value(lngYearIndex)
      If IsNull(varVal) Then
        strYear = ""
      Else
        strYear = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lngTurnIndex)
      If IsNull(varVal) Then
        strTurn = ""
      Else
        strTurn = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lngNotesIndex)
      If IsNull(varVal) Then
        strNotes = ""
      Else
        strNotes = Trim(CStr(varVal))
      End If
      varVal = pRow.Value(lngExtraNotesIndex)
      If IsNull(varVal) Then
       strExtra = ""
      Else
        strExtra = Trim(CStr(varVal))
      End If

      If Not MyGeneralOperations.CheckCollectionForKey(pReturn, strQuadrat) Then
        Set pQuadratByYearColl = ReturnEmptyYearColl
      Else
        Set pQuadratByYearColl = pReturn.Item(strQuadrat)
        pReturn.Remove strQuadrat
      End If

      If strYear <> "" And strTurn <> "" And strTurn <> "0" Then ' ONLY WORRY ABOUT CASES WHERE ROTATION IS DESIGNATED...
        varElement = Array(strSite, strPlot, strYear, strTurn, strNotes, strExtra)
        pQuadratByYearColl.Remove strYear
        pQuadratByYearColl.Add varElement, strYear
        Debug.Print "Plot '" & strPlot & "' [Quadrat = '" & strQuadrat & "'], " & strYear & ": Rotate " & strTurn
      End If

      pReturn.Add pQuadratByYearColl, strQuadrat
    End If

    Set pRow = pCursor.NextRow
  Loop

  Set FillRotateColl_SA = pReturn

ClearMemory:
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pMxDoc = Nothing
  Set pWSFact = Nothing
  Set pWS = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  Set pReturn = Nothing
  Erase varElement
  varVal = Null
  Set pQuadratByYearColl = Nothing

End Function

Public Function ReturnEmptyYearColl() As Collection
  Dim lngIndex As Long
  Dim pReturn As New Collection
  Dim varElement() As Variant

  For lngIndex = 1900 To 2140
    varElement = Array("", "", Format(lngIndex, "0"), "0", "", "")  ' PRESET ROTATION TO ZERO
    pReturn.Add varElement, Format(lngIndex, "0")
  Next lngIndex
  Set ReturnEmptyYearColl = pReturn

  Set pReturn = Nothing
  Erase varElement
End Function

Public Sub ShiftFinishedShapefilesToCoordinateSystem_SA()

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray

  Dim strOrigRoot As String
  Dim strModRoot As String
  Dim strShiftRoot As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, , strModRoot)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim strPlotLocNames() As String
  Dim pPlotLocColl As Collection

  Dim strPlotDataNames() As String
  Dim pPlotDataColl As Collection

  Dim strQuadratNames() As String
  Dim pQuadratColl As Collection
  Dim varSites() As Variant
  Dim varSiteSpecifics() As Variant

  Call ReturnQuadratVegSoilData(pPlotDataColl, strPlotDataNames)
  Call ReturnQuadratCoordsAndNames(pPlotLocColl, strPlotLocNames)
  Set pQuadratColl = FillQuadratNameColl_Rev_SA(strQuadratNames, , , varSites, varSiteSpecifics, booRestrictToSite, strSiteToRestrict)

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace
  Dim pCoverAll As IFeatureClass
  Dim pDensityAll As IFeatureClass
  Dim varCoverIndexes() As Variant
  Dim varDensityIndexes() As Variant

  Dim strFClassName As String
  Dim strNameSplit() As String

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pSrcWS = pNewWSFact.OpenFromFile(strModRoot & "\Combined_by_Site.gdb", 0)
  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strShiftRoot & "\Combined_by_Site")

  Set pWS = pSrcWS
  Set pDatasetEnum = pWS.Datasets(esriDTFeatureClass)
  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    If strFClassName <> "Cover_All" And strFClassName <> "Density_All" Then
      If InStr(1, strFClassName, "Density", vbTextCompare) Then
        Debug.Print strFClassName
      End If

      ExportFGDBFClass_2_SA pNewWS, pDataset, pMxDoc, pPlotLocColl, pQuadratColl, pCoverAll, pDensityAll, _
          varCoverIndexes, varDensityIndexes, False
    End If
    Set pDataset = pDatasetEnum.Next
  Loop

  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, CStr(varCoverIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Plot")

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, CStr(varDensityIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Plot")

  If Not aml_func_mod.ExistFileDir(strShiftRoot & "\Shapefiles") Then
    MyGeneralOperations.CreateNestedFoldersByPath (strShiftRoot & "\Shapefiles")
  End If
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strShiftRoot & "\Shapefiles", 0)

  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    If strFClassName <> "Cover_All" And strFClassName <> "Density_All" Then
      Debug.Print strFClassName

      ExportFGDBFClass_2_SA pNewWS, pDataset, pMxDoc, pPlotLocColl, pQuadratColl, pCoverAll, pDensityAll, _
          varCoverIndexes, varDensityIndexes, True
    End If
    Set pDataset = pDatasetEnum.Next
  Loop

  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, CStr(varCoverIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pCoverAll, "Plot")

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "SPCODE")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "FClassName")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Seedling")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Species")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Quadrat")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, CStr(varDensityIndexes(2, 9))) ' Year
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Type")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Orig_FID")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Verb_Spcs")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Site")
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDensityAll, "Plot")
  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Erase strPlotLocNames
  Set pPlotLocColl = Nothing
  Erase strPlotDataNames
  Set pPlotDataColl = Nothing
  Erase strQuadratNames
  Set pQuadratColl = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Set pCoverAll = Nothing
  Set pDensityAll = Nothing
  Erase varCoverIndexes
  Erase varDensityIndexes
  Erase strNameSplit

End Sub

Public Sub ExportFGDBFClass_2_SA(pDestWS As IFeatureWorkspace, pSrcFClass As IFeatureClass, _
    pMxDoc As IMxDocument, pPlotLocColl As Collection, pQuadratColl As Collection, pCoverAll As IFeatureClass, _
    pDensityAll As IFeatureClass, varCoverIndexes() As Variant, varDensityIndexes() As Variant, _
    booIsShapefile As Boolean)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngYearIndex As Long
  Dim pInsertFCursor As IFeatureCursor
  Dim pInsertFBuffer As IFeatureBuffer
  Dim pDestFClass As IFeatureClass
  Dim varIndexArray() As Variant
  Dim strNewName As String
  Dim lngIndex As Long
  Dim pDataset As IDataset
  Dim lngQuadratIndex As Long

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Dim pPolygon As IPolygon
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String

  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim lngCount As Long
  Dim lngCounter As Long
  lngCount = pSrcFClass.FeatureCount(Nothing)
  lngCounter = 0

  Set pDataset = pSrcFClass
  strNewName = pDataset.Name
  Debug.Print "  --> " & strNewName
  DoEvents
  If MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, strNewName) Then
    Set pDataset = pDestWS.OpenFeatureClass(strNewName)
    pDataset.DELETE
  End If

  Dim pDensityFCursor As IFeatureCursor
  Dim pDensityFBuffer As IFeatureBuffer
  Dim pCoverFCursor As IFeatureCursor
  Dim pCoverFBuffer As IFeatureBuffer
  Dim pClone As IClone
  Dim booDoCover As Boolean
  Dim booDoDensity As Boolean

  Set pDataset = pSrcFClass
  booDoCover = InStr(1, pDataset.BrowseName, "Cover", vbTextCompare)
  booDoDensity = InStr(1, pDataset.BrowseName, "Density", vbTextCompare)

  If booDoCover Then
    If Not MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, "Cover_All") Then
      Set pCoverAll = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varCoverIndexes, _
            "Cover_All", True)
      Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pCoverAll, strAbstract, strPurpose)
    Else
      For lngIndex = 0 To UBound(varCoverIndexes, 2)
        varCoverIndexes(1, lngIndex) = pSrcFClass.FindField(CStr(varCoverIndexes(0, lngIndex)))
      Next lngIndex
    End If
  End If

  If booDoDensity Then
    If Not MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, "Density_All") Then
      Set pDensityAll = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varDensityIndexes, _
            "Density_All", True)
      Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDensityAll, strAbstract, strPurpose)
    Else
      For lngIndex = 0 To UBound(varDensityIndexes, 2)
        varDensityIndexes(1, lngIndex) = pSrcFClass.FindField(CStr(varDensityIndexes(0, lngIndex)))
      Next lngIndex
    End If
  End If

  If booDoCover Then
    Set pCoverFCursor = pCoverAll.Insert(True)
    Set pCoverFBuffer = pCoverAll.CreateFeatureBuffer
  End If
  If booDoDensity Then
    Set pDensityFCursor = pDensityAll.Insert(True)
    Set pDensityFBuffer = pDensityAll.CreateFeatureBuffer
  End If

  Set pDestFClass = MyGeneralOperations.ReturnEmptyFClassWithSameSchema(pSrcFClass, pDestWS, varIndexArray, strNewName, True)
  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDestFClass, strAbstract, strPurpose)
  Set pInsertFCursor = pDestFClass.Insert(True)
  Set pInsertFBuffer = pDestFClass.CreateFeatureBuffer

  pSBar.ShowProgressBar "Exporting '" & pDataset.BrowseName & "'...", 0, lngCount, 1, True
  pProg.position = 0

  Dim dblCentroidX As Double
  Dim dblCentroidY As Double
  Dim strQuadrat As String
  Dim varItem() As Variant
  Dim strPlot As String

  Dim varSrcVal As Variant
  Dim lngVarType As Long
  Dim lngDestIndex As Long

  Dim lngSPCodeFieldIndex As Long
  lngSPCodeFieldIndex = pSrcFClass.FindField("SPCODE")

  lngQuadratIndex = pSrcFClass.FindField("Quadrat")

  Set pFCursor = pSrcFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
      pInsertFCursor.Flush
      If booDoCover Then pCoverFCursor.Flush
      If booDoDensity Then pDensityFCursor.Flush
    End If

    strQuadrat = pFeature.Value(lngQuadratIndex)
    varItem = pQuadratColl.Item(Replace(strQuadrat, "Q", ""))
    strPlot = varItem(2)
    FillQuadratCenter_SA strPlot, pQuadratColl, dblCentroidX, dblCentroidY

    Set pPolygon = pFeature.ShapeCopy
    Call Margaret_Functions.ShiftPolygon(pPolygon, dblCentroidX, dblCentroidY)
    Set pClone = pPolygon

    Set pInsertFBuffer.Shape = pClone.Clone
    For lngIndex = 0 To UBound(varIndexArray, 2)

      varSrcVal = pFeature.Value(varIndexArray(1, lngIndex))
      lngDestIndex = varIndexArray(3, lngIndex)
      If booIsShapefile Then
        If IsNull(varSrcVal) Then
          If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
            varSrcVal = ""
          End If
        End If
      End If
      pInsertFBuffer.Value(lngDestIndex) = varSrcVal
    Next lngIndex
    pInsertFCursor.InsertFeature pInsertFBuffer

    If booDoDensity Then
      Set pDensityFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varDensityIndexes, 2)
        varSrcVal = pFeature.Value(varDensityIndexes(1, lngIndex))
        lngDestIndex = varDensityIndexes(3, lngIndex)

        If varDensityIndexes(1, lngIndex) = lngSPCodeFieldIndex Then ' if SPCODE field, which should be integer
          If IsNull(varSrcVal) Then
            If booIsShapefile Then
              pDensityFBuffer.Value(lngDestIndex) = -999
            Else
              pDensityFBuffer.Value(lngDestIndex) = Null
            End If
          Else
            If Trim(CStr(pFeature.Value(varDensityIndexes(1, lngIndex)))) = "" Then
              pDensityFBuffer.Value(lngDestIndex) = Null
            Else
              pDensityFBuffer.Value(lngDestIndex) = pFeature.Value(varDensityIndexes(1, lngIndex))
            End If
          End If
        Else

          If booIsShapefile Then
            If IsNull(varSrcVal) Then
              If pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
                varSrcVal = -999
              ElseIf pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
                varSrcVal = -999
              ElseIf pDensityFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
                varSrcVal = ""
              End If
            End If
            pDensityFBuffer.Value(lngDestIndex) = varSrcVal
          Else
            pDensityFBuffer.Value(varDensityIndexes(3, lngIndex)) = pFeature.Value(varDensityIndexes(1, lngIndex))
          End If

        End If
      Next lngIndex
      pDensityFCursor.InsertFeature pDensityFBuffer
    End If

    If booDoCover Then
      Set pCoverFBuffer.Shape = pClone.Clone
      For lngIndex = 0 To UBound(varCoverIndexes, 2)
        varSrcVal = pFeature.Value(varCoverIndexes(1, lngIndex))
        lngDestIndex = varCoverIndexes(3, lngIndex)

        If varCoverIndexes(1, lngIndex) = lngSPCodeFieldIndex Then ' if SPCODE field, which should be integer
          If IsNull(varSrcVal) Then
            If booIsShapefile Then
              pCoverFBuffer.Value(lngDestIndex) = -999
            Else
              pCoverFBuffer.Value(lngDestIndex) = Null
            End If
          Else
            If Trim(CStr(pFeature.Value(varCoverIndexes(1, lngIndex)))) = "" Then
              pCoverFBuffer.Value(lngDestIndex) = Null
            Else
              pCoverFBuffer.Value(lngDestIndex) = pFeature.Value(varCoverIndexes(1, lngIndex))
            End If
          End If

        Else
          If booIsShapefile Then
            lngDestIndex = varCoverIndexes(3, lngIndex)
            varSrcVal = pFeature.Value(varCoverIndexes(1, lngIndex))
            If IsNull(varSrcVal) Then
              If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
                varSrcVal = -999
              ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
                varSrcVal = -999
              ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
                varSrcVal = ""
              End If
            End If
            pCoverFBuffer.Value(lngDestIndex) = varSrcVal
          Else
            pCoverFBuffer.Value(varCoverIndexes(3, lngIndex)) = pFeature.Value(varCoverIndexes(1, lngIndex))
          End If
        End If
      Next lngIndex
      pCoverFCursor.InsertFeature pCoverFBuffer
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pInsertFCursor.Flush
  If booDoCover Then pCoverFCursor.Flush
  If booDoDensity Then pDensityFCursor.Flush

  pSBar.ShowProgressBar "Building Indexes for '" & pDataset.BrowseName & "'...", 0, 8, 1, True
  pProg.position = 0

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "SPCODE")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "FClassName")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Seedling")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Species")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Quadrat")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, CStr(varIndexArray(2, 9))) ' Year
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Orig_FID")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Site")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Verb_Spcs")
  pProg.Step
  DoEvents

  pSBar.HideProgressBar
  pProg.position = 0

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pInsertFCursor = Nothing
  Set pInsertFBuffer = Nothing
  Set pDestFClass = Nothing
  Erase varIndexArray
  Set pDataset = Nothing
  Set pPolygon = Nothing
  Set pQueryFilt = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pDensityFCursor = Nothing
  Set pDensityFBuffer = Nothing
  Set pCoverFCursor = Nothing
  Set pCoverFBuffer = Nothing
  Set pClone = Nothing

End Sub

Public Sub ExportFinalFClass(pDestWS As IFeatureWorkspace, pSrcFClass As IFeatureClass, _
    pMxDoc As IMxDocument, booIsShapefile As Boolean, pVegColl As Collection)

  Dim varVegData() As Variant

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngYearIndex As Long
  Dim pInsertFCursor As IFeatureCursor
  Dim pInsertFBuffer As IFeatureBuffer
  Dim pDestFClass As IFeatureClass
  Dim varIndexArray() As Variant
  Dim strNewName As String
  Dim lngIndex As Long
  Dim pDataset As IDataset
  Dim lngQuadratIndex As Long

  Dim strAbstract As String
  Dim strBaseString As String
  Dim strPurpose As String
  More_Margaret_Functions.FillMetadataItems strAbstract, strBaseString, strPurpose

  Dim pPolygon As IPolygon
  Dim pQueryFilt As IQueryFilter
  Dim strPrefix As String
  Dim strSuffix As String

  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim lngCount As Long
  Dim lngCounter As Long
  lngCount = pSrcFClass.FeatureCount(Nothing)
  lngCounter = 0

  Set pDataset = pSrcFClass
  strNewName = pDataset.Name
  Debug.Print "  --> " & strNewName
  DoEvents
  If MyGeneralOperations.CheckIfFeatureClassExists(pDestWS, strNewName) Then
    Set pDataset = pDestWS.OpenFeatureClass(strNewName)
    pDataset.DELETE
  End If

  Dim pClone As IClone

  Set pDataset = pSrcFClass

  Dim lngSPCodeFieldIndex As Long
  Dim pNewField As iField
  lngSPCodeFieldIndex = pSrcFClass.FindField("SPCODE")
  If pSrcFClass.Fields.Field(lngSPCodeFieldIndex).Type <> esriFieldTypeString Then
    pSrcFClass.DeleteField pSrcFClass.Fields.Field(lngSPCodeFieldIndex)
    Set pNewField = MyGeneralOperations.CreateNewField("SPCODE", esriFieldTypeString, , 10)
    pSrcFClass.AddField pNewField
    lngSPCodeFieldIndex = pSrcFClass.FindField("SPCODE")
  End If

  Set pDestFClass = ReturnEmptyFClassWithSameSchema_SpecialCase(pSrcFClass, pDestWS, varIndexArray, strNewName, True)
  Call Margaret_Functions.Metadata_pNewFClass(pMxDoc, pDestFClass, strAbstract, strPurpose)

  lngSPCodeFieldIndex = pDestFClass.FindField("SPCODE")

  Set pInsertFCursor = pDestFClass.Insert(True)
  Set pInsertFBuffer = pDestFClass.CreateFeatureBuffer

  pSBar.ShowProgressBar "Exporting '" & pDataset.BrowseName & "'...", 0, lngCount, 1, True
  pProg.position = 0

  Dim dblCentroidX As Double
  Dim dblCentroidY As Double
  Dim strQuadrat As String
  Dim strItem() As String
  Dim strPlot As String

  Dim varSrcVal As Variant
  Dim lngVarType As Long
  Dim lngDestIndex As Long
  Dim strSpecies As String
  Dim strSPCode As String
  Dim lngSpeciesIndex As Long

  lngSpeciesIndex = pSrcFClass.FindField("Species")
  lngQuadratIndex = pSrcFClass.FindField("Quadrat")

  Set pFCursor = pSrcFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
      pInsertFCursor.Flush
    End If

    strSpecies = pFeature.Value(lngSpeciesIndex)
    If MyGeneralOperations.CheckCollectionForKey(pVegColl, strSpecies) Then
      varVegData = pVegColl.Item(strSpecies)
      strSPCode = varVegData(9)
    Else
      strSPCode = ""
    End If

    Set pInsertFBuffer.Shape = pFeature.ShapeCopy ' pClone.Clone
    For lngIndex = 0 To UBound(varIndexArray, 2)

      varSrcVal = pFeature.Value(varIndexArray(1, lngIndex))
      lngDestIndex = varIndexArray(3, lngIndex)
      If booIsShapefile Then
        If IsNull(varSrcVal) Then
          If pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeInteger Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeDouble Then
            varSrcVal = -999
          ElseIf pInsertFBuffer.Fields.Field(lngDestIndex).Type = esriFieldTypeString Then
            varSrcVal = ""
          End If
        End If
      End If
      pInsertFBuffer.Value(lngDestIndex) = varSrcVal
    Next lngIndex
    pInsertFBuffer.Value(lngSPCodeFieldIndex) = strSPCode

    pInsertFCursor.InsertFeature pInsertFBuffer

    Set pFeature = pFCursor.NextFeature
  Loop

  pInsertFCursor.Flush

  pSBar.ShowProgressBar "Building Indexes for '" & pDataset.BrowseName & "'...", 0, 8, 1, True
  pProg.position = 0

  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "SPCODE")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "FClassName")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Seedling")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Species")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Quadrat")
  pProg.Step
  DoEvents
  If pDestFClass.FindField("Year") > -1 Then
    Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Year") ' Year
  Else
    Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "z_Year") ' Year
  End If
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Orig_FID")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Site")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Type")
  pProg.Step
  DoEvents
  Call MyGeneralOperations.CreateFieldAttributeIndex(pDestFClass, "Verb_Spcs")
  pProg.Step
  DoEvents

  pSBar.HideProgressBar
  pProg.position = 0

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pInsertFCursor = Nothing
  Set pInsertFBuffer = Nothing
  Set pDestFClass = Nothing
  Erase varIndexArray
  Set pDataset = Nothing
  Set pPolygon = Nothing
  Set pQueryFilt = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pClone = Nothing
  Erase strItem
  varSrcVal = Null

End Sub

Public Function ReturnEmptyFClassWithSameSchema_SpecialCase(pFClass As IFeatureClass, pWS_NothingForInMemory As IWorkspace, _
    varFieldIndexArray() As Variant, strName As String, booHasFields As Boolean, _
    Optional lngForceGeometryType As esriGeometryType = esriGeometryAny) As IFeatureClass

  Dim pFields As IFields
  Set pFields = pFClass.Fields
  Dim booIsShapefile As Boolean
  Dim booIsAccess As Boolean
  Dim booIsFGDB As Boolean
  Dim booIsInMem As Boolean
  Dim lngCategory As JenDatasetTypes

  If lngForceGeometryType = esriGeometryAny Then
    lngForceGeometryType = pFClass.ShapeType
  End If

  If pWS_NothingForInMemory Is Nothing Then
    booIsInMem = True
  Else
    booIsInMem = False
    booIsShapefile = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Shapefile Workspace Factory"
    booIsAccess = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "Esri Access Workspace Factory"
    booIsFGDB = ReturnWorkspaceFactoryType(pWS_NothingForInMemory.WorkspaceFactory.GetClassID) = "File GeoDatabase Workspace Factory"
  End If

  If booIsAccess Then
    lngCategory = ENUM_PersonalGDB
  ElseIf booIsFGDB Then
    lngCategory = ENUM_FileGDB
  End If

  Dim lngIndex As Long
  Dim pSrcField As iField
  Dim pNewField As iField
  Dim pNewFieldEdit As IFieldEdit
  Dim pClone As IClone

  Dim pNewFieldArray As esriSystem.IVariantArray
  Set pNewFieldArray = New esriSystem.varArray

  Dim lngCounter As Long
  lngCounter = -1
  Dim varReturnArray() As Variant

  For lngIndex = 0 To pFields.FieldCount - 1
    Set pSrcField = pFields.Field(lngIndex)
    If Not pSrcField.Type = esriFieldTypeOID And pSrcField.Type <> esriFieldTypeGeometry And _
        StrComp(Left(pSrcField.Name, 6), "Shape_", vbTextCompare) <> 0 Then
      If pSrcField.Name <> "Plot" And pSrcField.Name <> "Verb_Spcs" And pSrcField.Name <> "Verb_Type" And _
           pSrcField.Name <> "Revise_Rtn" And pSrcField.Name <> "FClassName" And pSrcField.Name <> "Orig_FID" Then
        Set pClone = pSrcField
        Set pNewField = pClone.Clone
        Set pNewFieldEdit = pNewField
        With pNewFieldEdit
          If booIsShapefile Then
            .Name = MyGeneralOperations.ReturnAcceptableFieldName2(pSrcField.Name, pNewFieldArray, booIsShapefile, booIsAccess, False, booIsFGDB)
            .IsNullable = False
            If pSrcField.Type = esriFieldTypeDouble Then
              .Precision = 16
              .Scale = 6
            End If
          Else
            .IsNullable = True
          End If
          If pSrcField.Name = "Quadrat" Then
            .length = 25
          End If
        End With
        pNewFieldArray.Add pNewField

        lngCounter = lngCounter + 1
        ReDim Preserve varReturnArray(3, lngCounter)
        If pSrcField.Name = "Quadrat" Then
          varReturnArray(0, lngCounter) = "Plot"
          varReturnArray(1, lngCounter) = pFields.FindField("Plot")
          varReturnArray(2, lngCounter) = pNewField.Name
        Else
          varReturnArray(0, lngCounter) = pSrcField.Name
          varReturnArray(1, lngCounter) = lngIndex
          varReturnArray(2, lngCounter) = pNewField.Name
        End If
      End If

    End If
  Next lngIndex

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Set pDataset = pFClass
  Set pGeoDataset = pFClass
  Dim pGeomDef As IGeometryDef
  Set pGeomDef = pFClass.Fields.Field(pFClass.FindField(pFClass.ShapeFieldName)).GeometryDef

  Dim pNewFClass As IFeatureClass

  Dim pEnv As IEnvelope
  Set pEnv = New Envelope
  pEnv.PutCoords -5, -5, 5, 5
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)

  If booIsInMem Then
    Set pNewFClass = MyGeneralOperations.CreateInMemoryFeatureClass_Empty(pNewFieldArray, strName, pSpRef, _
        lngForceGeometryType, pGeomDef.HasM, pGeomDef.HasZ)
  ElseIf booIsFGDB Or booIsAccess Then
    Set pNewFClass = MyGeneralOperations.CreateGDBFeatureClass2(pWS_NothingForInMemory, strName, esriFTSimple, pSpRef, _
        lngForceGeometryType, pNewFieldArray, , , , False, lngCategory, pEnv, , pGeomDef.HasZ, pGeomDef.HasM)
  ElseIf booIsShapefile Then
    Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(pWS_NothingForInMemory.PathName, strName, _
        pSpRef, lngForceGeometryType, pNewFieldArray, False, pGeomDef.HasZ, pGeomDef.HasM)
  Else
    MsgBox "No code written for this workspace type!"
    GoTo ClearMemory
  End If

  booHasFields = lngCounter > -1
  If booHasFields Then
    For lngIndex = 0 To lngCounter
      varReturnArray(3, lngIndex) = pNewFClass.FindField(CStr(varReturnArray(2, lngIndex)))
    Next lngIndex
  End If
  varFieldIndexArray = varReturnArray
  Set ReturnEmptyFClassWithSameSchema_SpecialCase = pNewFClass

  GoTo ClearMemory
ClearMemory:
  Set pFields = Nothing
  Set pSrcField = Nothing
  Set pNewField = Nothing
  Set pNewFieldEdit = Nothing
  Set pClone = Nothing
  Set pNewFieldArray = Nothing
  Erase varReturnArray
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pGeomDef = Nothing
  Set pNewFClass = Nothing

End Function

Public Function ReturnArrayOfFieldLinks(pSrcFClass As IFeatureClass, pDestFClass As IFeatureClass) As Variant()

  Dim pSrcFields As IFields
  Dim pDestFields As IFields
  Set pSrcFields = pSrcFClass.Fields
  Set pDestFields = pDestFClass.Fields

  Dim pField As iField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String

  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry Then
      lngCounter = lngCounter + 1
      ReDim Preserve varReturn(3, lngCounter)
      varReturn(0, lngCounter) = pField.Name
      varReturn(1, lngCounter) = lngIndex
      If pField.Type = esriFieldTypeOID Then
        strNewName = "Orig_FID"
      ElseIf pField.Name = "SP_CODE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP_CPDE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP_" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP" Then
        strNewName = "SPCODE"
      Else
        strNewName = pField.Name
      End If
      varReturn(2, lngCounter) = strNewName
      varReturn(3, lngCounter) = pDestFields.FindField(strNewName)

          " | " & varReturn(2, lngCounter) & " | " & varReturn(3, lngCounter)

    End If
  Next lngIndex

  ReturnArrayOfFieldLinks = varReturn

  Set pSrcFields = Nothing
  Set pDestFields = Nothing
  Erase varReturn

End Function

Public Function ReturnArrayOfFieldCrossLinks(pSrcFClass As IFeatureClass, pDestFClass As IFeatureClass) As Variant()

  Dim pSrcFields As IFields
  Dim pDestFields As IFields
  Set pSrcFields = pSrcFClass.Fields
  Set pDestFields = pDestFClass.Fields

  Dim pField As iField
  Dim lngIndex As Long
  Dim varReturn() As Variant
  Dim lngCounter As Long
  Dim strNewName As String

  lngCounter = -1
  For lngIndex = 0 To pSrcFields.FieldCount - 1
    Set pField = pSrcFields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry Then
      If pField.Type = esriFieldTypeOID Then
        strNewName = "Orig_FID"
      ElseIf pField.Name = "SP_CODE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP_CPDE" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP_" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SPP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "SP" Then
        strNewName = "SPCODE"
      ElseIf pField.Name = "x" Then
        strNewName = "coords_x1"
      ElseIf pField.Name = "y" Then
        strNewName = "coords_x2"
      ElseIf pField.Name = "coords_x1" Then
        strNewName = "x"
      ElseIf pField.Name = "coords_x2" Then
        strNewName = "y"
      Else
        strNewName = pField.Name
      End If

      If pDestFields.FindField(strNewName) > -1 Then
        lngCounter = lngCounter + 1
        ReDim Preserve varReturn(3, lngCounter)
        varReturn(0, lngCounter) = pField.Name
        varReturn(1, lngCounter) = lngIndex
        varReturn(2, lngCounter) = strNewName
        varReturn(3, lngCounter) = pDestFields.FindField(strNewName)
      End If

          " | " & varReturn(2, lngCounter) & " | " & varReturn(3, lngCounter)

    End If
  Next lngIndex

  ReturnArrayOfFieldCrossLinks = varReturn

  Set pSrcFields = Nothing
  Set pDestFields = Nothing
  Erase varReturn

End Function

Public Sub CreateNewFields(pNewFClass As IFeatureClass, lngFClassNameIndex As Long, _
    lngQuadratIndex As Long, lngYearIndex As Long, lngTypeIndex As Long, lngOrigFIDIndex As Long)

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim lngIDIndex As Long
  Dim lngIsEmptyIndex As Long

  Dim lngSPCodeIndex As Long
  lngSPCodeIndex = pNewFClass.FindField("SP_CODE")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SP_CPDE")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPP_")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPP")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SP")
  If lngSPCodeIndex > -1 Then
    Set pField = pNewFClass.Fields.Field(lngSPCodeIndex)
    pNewFClass.DeleteField pField
  End If
  lngSPCodeIndex = pNewFClass.FindField("SPCODE")
  If lngSPCodeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "SPCODE"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngSPCodeIndex = pNewFClass.FindField("SPCODE")
  End If

  lngIDIndex = pNewFClass.FindField("Id")
  If lngIDIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Id"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngIDIndex = pNewFClass.FindField("Id")
  End If

  lngFClassNameIndex = pNewFClass.FindField("FClassName")
  If lngFClassNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "FClassName"
      .Type = esriFieldTypeString
      .length = 50
    End With
    pNewFClass.AddField pField
    lngFClassNameIndex = pNewFClass.FindField("FClassName")
  End If

  lngQuadratIndex = pNewFClass.FindField("Quadrat")
  If lngQuadratIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Quadrat"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngQuadratIndex = pNewFClass.FindField("Quadrat")
  End If

  lngYearIndex = pNewFClass.FindField("Year")
  If lngYearIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Year"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngYearIndex = pNewFClass.FindField("Year")
  End If

  lngTypeIndex = pNewFClass.FindField("Type")
  If lngTypeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Type"
      .Type = esriFieldTypeString
      .length = 10
    End With
    pNewFClass.AddField pField
    lngTypeIndex = pNewFClass.FindField("Type")
  End If

  lngOrigFIDIndex = pNewFClass.FindField("Orig_FID")
  If lngOrigFIDIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Orig_FID"
      .Type = esriFieldTypeInteger
    End With
    pNewFClass.AddField pField
    lngOrigFIDIndex = pNewFClass.FindField("Orig_FID")
  End If

  lngIsEmptyIndex = pNewFClass.FindField("IsEmpty")
  If lngIsEmptyIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "IsEmpty"
      .Type = esriFieldTypeString
      .length = 5
    End With
    pNewFClass.AddField pField
    lngIsEmptyIndex = pNewFClass.FindField("IsEmpty")
  End If

  Set pField = Nothing
  Set pFieldEdit = Nothing

End Sub

Public Function CheckSpeciesAgainstSpecialConversions(varSpecialConversions() As Variant, strQuadrat As String, _
    lngYear As Long, strSpecies As String, strNoteOnChanges As String) As String

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long
  Dim strTestSpecies As String

  If InStr(1, strSpecies, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
    DoEvents
  End If

  CheckSpeciesAgainstSpecialConversions = Trim(strSpecies)
  strNoteOnChanges = ""

  For lngIndex = 0 To UBound(varSpecialConversions, 2)
    strTestQuadrat = varSpecialConversions(0, lngIndex)
    lngTestYear = varSpecialConversions(1, lngIndex)
    strTestSpecies = varSpecialConversions(2, lngIndex)
    If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
      If lngTestYear = lngYear Or lngTestYear = -999 Then
        If StrComp(Trim(strSpecies), Trim(strTestSpecies), vbTextCompare) = 0 Then
          CheckSpeciesAgainstSpecialConversions = Trim(CStr(varSpecialConversions(3, lngIndex)))
          strNoteOnChanges = Trim(CStr(varSpecialConversions(5, lngIndex)))
          Exit Function
        End If
      End If
    End If
  Next lngIndex

End Function

Public Function SpecialConversionExistsForYearQuadrat(varSpecialConversions() As Variant, strQuadrat As String, _
    lngYear As Long) As Boolean

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long

  SpecialConversionExistsForYearQuadrat = False

  For lngIndex = 0 To UBound(varSpecialConversions, 2)
    strTestQuadrat = varSpecialConversions(0, lngIndex)
    lngTestYear = varSpecialConversions(1, lngIndex)
    If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
      If lngTestYear = lngYear Or lngTestYear = -999 Then
        SpecialConversionExistsForYearQuadrat = True
        Exit Function
      End If
    End If
  Next lngIndex

End Function

Public Function ReturnQueryStringFromSpecialConversions(strQuadrat As String, lngYear As Long, _
    booIsCover As Boolean, varSpecialConversions() As Variant, strInstructions As String, _
    lngSpecialIndex As Long, booYearQuadrat As Boolean) As String

  Dim lngIndex As Long
  Dim strTestQuadrat As String
  Dim lngTestYear As Long
  Dim strTestSpecies As String

  Dim strQueryString As String
  Dim varStrings() As Variant
  Dim varInstructions() As Variant

  ReturnQueryStringFromSpecialConversions = ""
  strInstructions = ""

  booYearQuadrat = False

  lngIndex = lngSpecialIndex
  strTestQuadrat = varSpecialConversions(0, lngIndex)
  lngTestYear = varSpecialConversions(1, lngIndex)
  If StrComp(Trim(strQuadrat), Trim(strTestQuadrat), vbTextCompare) = 0 Then
    If lngTestYear = lngYear Then
      booYearQuadrat = True
      varStrings = varSpecialConversions(4, lngIndex)
      varInstructions = varSpecialConversions(6, lngIndex)
      If booIsCover Then
        ReturnQueryStringFromSpecialConversions = varStrings(0)
        strInstructions = varInstructions(0)
      Else
        ReturnQueryStringFromSpecialConversions = varStrings(1)
        strInstructions = varInstructions(1)
      End If
      Exit Function
    End If
  End If

ClearMemory:
  Erase varStrings
  Erase varInstructions

End Function

Public Sub CopyFeaturesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, _
    strQueryPair As String, strEditReport As String, strExcelReport As String, _
    booMadeEdits As Boolean, lngNameIndex As Long, strBase As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strOID As String

  Dim strSource As String
  Dim strQueryString As String
  Dim strQuerySplit() As String
  Dim pQueryFilt As IQueryFilter

  strQuerySplit = Split(strQueryPair, "|")
  strSource = strQuerySplit(0)
  strQueryString = strQuerySplit(1)
  Set pQueryFilt = New QueryFilter
  pQueryFilt.WhereClause = strQueryString

  Dim pWS As IFeatureWorkspace
  Dim pDataset As IDataset
  Dim pDonorFClass As IFeatureClass

  Set pDataset = pFClass
  Set pWS = pDataset.Workspace
  Set pDonorFClass = pWS.OpenFeatureClass(strSource)

  Dim lngLinks() As Long
  Dim lngLinkIndex As Long
  Dim pField As iField
  Dim lngIndex As Long

  lngLinkIndex = -1
  For lngIndex = 0 To pDonorFClass.Fields.FieldCount - 1
    Set pField = pDonorFClass.Fields.Field(lngIndex)
    If pField.Type <> esriFieldTypeGeometry And pField.Type <> esriFieldTypeOID Then
      If pFClass.FindField(pField.Name) > -1 Then
        lngLinkIndex = lngLinkIndex + 1
        ReDim Preserve lngLinks(1, lngLinkIndex)
        lngLinks(0, lngLinkIndex) = lngIndex  ' DONOR INDEX
        lngLinks(1, lngLinkIndex) = pFClass.FindField(pField.Name) ' RECIPIENT INDEX
      End If
    End If
  Next lngIndex

  Dim varFeatures() As Variant
  Dim lngArrayIndex As Long

  lngArrayIndex = -1
  Set pFCursor = pDonorFClass.Search(pQueryFilt, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    lngArrayIndex = lngArrayIndex + 1
    ReDim Preserve varFeatures(lngArrayIndex)
    Set varFeatures(lngArrayIndex) = pFeature
    Set pFeature = pFCursor.NextFeature
  Loop

  Dim lngSrcSpeciesIndex As Long
  Dim lngDestSpeciesIndex As Long
  Dim strSrcSpecies As String
  Dim strDestSpecies As String
  Dim pSrcEnv As IEnvelope
  Dim pDestEnv As IEnvelope
  lngSrcSpeciesIndex = pDonorFClass.FindField("species")
  lngDestSpeciesIndex = pFClass.FindField("species")

  Dim varDestData() As Variant
  Dim lngDestIndex As Long
  lngDestIndex = -1
  Set pFCursor = pFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pDestEnv = pFeature.ShapeCopy.Envelope
    strDestSpecies = pFeature.Value(lngDestSpeciesIndex)

    lngDestIndex = lngDestIndex + 1
    ReDim Preserve varDestData(1, lngDestIndex)
    varDestData(0, lngDestIndex) = strDestSpecies
    Set varDestData(1, lngDestIndex) = pDestEnv

    Set pFeature = pFCursor.NextFeature
  Loop

  Dim pFBuffer As IFeatureBuffer
  Dim booFeatureExists As Boolean
  Dim lngIndex2 As Long

  If lngArrayIndex > -1 Then

    Set pFBuffer = pFClass.CreateFeatureBuffer
    Set pFCursor = pFClass.Insert(True)
    For lngIndex = 0 To lngArrayIndex
      Set pFeature = varFeatures(lngIndex)
      Set pSrcEnv = pFeature.ShapeCopy.Envelope
      strSrcSpecies = pFeature.Value(lngSrcSpeciesIndex)

      booFeatureExists = False
      If lngDestIndex > -1 Then
        For lngIndex2 = 0 To lngDestIndex
          strDestSpecies = varDestData(0, lngIndex2)
          If StrComp(Trim(strDestSpecies), Trim(strSrcSpecies), vbTextCompare) = 0 Then
            Set pDestEnv = varDestData(1, lngIndex2)
            If pDestEnv.XMin = pSrcEnv.XMin And pDestEnv.XMax = pSrcEnv.XMax And _
                    pDestEnv.YMin = pSrcEnv.YMin And pDestEnv.YMax = pSrcEnv.YMax Then
              booFeatureExists = True
              Exit For
            End If
          End If
        Next lngIndex2
      End If

      If Not booFeatureExists Then
        Set pFBuffer.Shape = pFeature.ShapeCopy
        For lngIndex2 = 0 To UBound(lngLinks, 2)
          pFBuffer.Value(lngLinks(1, lngIndex2)) = pFeature.Value(lngLinks(0, lngIndex2))
        Next lngIndex2
        pFCursor.InsertFeature pFBuffer
        pFCursor.Flush

        booMadeEdits = True
        strOID = CStr(pFBuffer.Value(pFBuffer.Fields.FindField(pFClass.OIDFieldName)))
        strOID = String(4 - Len(strOID), " ") & strOID

        strName = pFBuffer.Value(lngNameIndex)
        strOrigName = strName

        strName = Replace(strName, vbCrLf, "")
        strName = Replace(strName, vbNewLine, "")
        strName = Trim(strName)
        Debug.Print "  --> " & CStr(pFeature.OID) & "] Copying new " & strName & " from " & strSource & "..."

        strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Species '" & _
            strName & "': Feature copied from " & strSource & " on Where Clause = " & _
            strQueryString & vbCrLf
        strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(strOID) & """" & vbTab & _
              """" & strName & """" & vbTab & """Feature copied from " & strSource & """" & vbCrLf

      End If
    Next lngIndex
  End If

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strQuerySplit
  Set pQueryFilt = Nothing
  Set pWS = Nothing
  Set pDataset = Nothing
  Set pDonorFClass = Nothing
  Erase lngLinks
  Set pField = Nothing
  Erase varFeatures
  Set pSrcEnv = Nothing
  Set pDestEnv = Nothing
  Erase varDestData
  Set pFBuffer = Nothing

End Sub

Public Sub DeleteFeaturesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, _
    pQueryFiltOrNothing As IQueryFilter, strEditReport As String, strExcelReport As String, _
    booMadeEdits As Boolean, lngNameIndex As Long, strBase As String)

  Dim pTable As ITable
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strOID As String

  Set pTable = pFClass
  If pTable.RowCount(pQueryFiltOrNothing) > 0 Then

    Set pFCursor = pFClass.Search(pQueryFiltOrNothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing

      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID

      strName = pFeature.Value(lngNameIndex)
      strOrigName = strName

      strName = Replace(strName, vbCrLf, "")
      strName = Replace(strName, vbNewLine, "")
      strName = Trim(strName)
      Debug.Print "  --> " & CStr(pFeature.OID) & "] Deleting " & strName & "..."

      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Species '" & _
          strName & "': Feature Deleted using 'DeleteSearchedRows' method on Where Clause = " & _
          pQueryFiltOrNothing.WhereClause & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """Feature Deleted""" & vbCrLf

      Set pFeature = pFCursor.NextFeature
    Loop

    pTable.DeleteSearchedRows pQueryFiltOrNothing
  End If

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pTable = Nothing

End Sub

Public Sub UpdateGeometryInFClassBasedOnQueryFilter(pFClass As IFeatureClass, pQueryFiltOrNothing As IQueryFilter, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String, strSourceSpecies As String, strDestSpecies As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Dim pOrigPolygon As IPolygon
  Dim pNewPolygon As IPolygon
  Dim pCentroid As IPoint
  Dim pArea As IArea

  Set pFCursor = pFClass.Update(pQueryFiltOrNothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    Set pOrigPolygon = pFeature.ShapeCopy
    Set pArea = pOrigPolygon
    Set pCentroid = pArea.Centroid
    Set pNewPolygon = MyGeometricOperations.CreateCircleAroundPoint(pCentroid, 0.001, 36)
    Set pFeature.Shape = pNewPolygon
    pFCursor.UpdateFeature pFeature

    booMadeEdits = True
    strOID = CStr(pFeature.OID)
    strOID = String(4 - Len(strOID), " ") & strOID
    strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed Polygon Geometry to Centroid'" & vbCrLf
    strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
          """" & strName & """" & vbTab & """Converted to Polygon Centroid""" & vbCrLf

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub UpdateSpeciesInFClassBasedOnQueryFilter(pFClass As IFeatureClass, pQueryFiltOrNothing As IQueryFilter, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String, strSourceSpecies As String, strDestSpecies As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Set pFCursor = pFClass.Update(pQueryFiltOrNothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    strOrigName = strName

    strName = Replace(strName, vbCrLf, "")
    strName = Replace(strName, vbNewLine, "")
    strName = Trim(strName)

    If InStr(1, strName, "tricholepis", vbTextCompare) > 0 Then
      DoEvents
    End If

    strCorrect = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strName, strNoteOnChanges)

    strHexify = HexifyName(strName)
    If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
      strCorrect = pCheckCollection.Item(strHexify)
    End If

    strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
    strCorrect = Replace(strCorrect, "Pachera ", "Packera ")
    If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
      MsgBox "Found carriage return!"
      DoEvents
    End If

    If InStr(1, strCorrect, " Asclepias sp.", vbTextCompare) > 0 Then
      DoEvents
    End If

    If InStr(1, strCorrect, "formossisimus", vbTextCompare) > 0 Then
      DoEvents
      strCorrect = Replace(strCorrect, "Erigeron formossisimus", "Erigeron formosissimus")
    End If

    If Left(strCorrect, 1) = " " Then
      strTrimName = Trim(strCorrect)
      If strCorrect <> " " Then
        DoEvents
      End If
      strCorrect = strTrimName
      strHexify = HexifyName(strTrimName)
      If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
        strCorrect = pCheckCollection.Item(strHexify)
      End If
      strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
      If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
        MsgBox "Found carriage return!"
        DoEvents
      End If
    End If

    If Left(strCorrect, 1) = " " Or Left(strName, 1) = " " Or InStr(1, strCorrect, vbTab) > 0 Or InStr(1, strName, vbTab) > 0 Then
      If strName <> " " Then
        DoEvents
      End If
    End If

    If StrComp(Trim(strOrigName), Trim(strSourceSpecies), vbTextCompare) = 0 Or _
       StrComp(Trim(strName), Trim(strSourceSpecies), vbTextCompare) = 0 Or _
       StrComp(Trim(strCorrect), Trim(strSourceSpecies), vbTextCompare) = 0 Then

      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID
      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed '" & _
          strName & "' to '" & strDestSpecies & "'" & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """" & strDestSpecies & """" & vbCrLf
      pFeature.Value(lngNameIndex) = strDestSpecies
      pFCursor.UpdateFeature pFeature
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub UpdateSpeciesInFClassBasedConversionArray(pFClass As IFeatureClass, _
    varSpecialConversions() As Variant, strQuadrat As String, strYear As String, strEditReport As String, _
    strExcelReport As String, booMadeEdits As Boolean, lngNameIndex As Long, pCheckCollection As Collection, _
    strBase As String)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim strName As String
  Dim strOrigName As String
  Dim strCorrect As String
  Dim strHexify As String
  Dim strTrimName As String
  Dim strOID As String
  Dim strNoteOnChanges As String

  Set pFCursor = pFClass.Update(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    strName = pFeature.Value(lngNameIndex)
    strOrigName = strName

    strName = Replace(strName, vbCrLf, "")
    strName = Replace(strName, Chr(9), "")  ' PROBLEM INTRODUCED 2020 WITH Q30 COVER Potentilla crinita
    strName = Replace(strName, vbNewLine, "")
    strName = Trim(strName)

    If InStr(1, strName, "Gutierrezia", vbTextCompare) > 0 Then
      DoEvents
    End If

    strCorrect = CheckSpeciesAgainstSpecialConversions(varSpecialConversions, strQuadrat, CLng(strYear), _
                strName, strNoteOnChanges)

    strHexify = HexifyName(strName)
    If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
      strCorrect = pCheckCollection.Item(strHexify)
    End If

    strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
    strCorrect = Replace(strCorrect, "Pachera ", "Packera ")
    If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
      MsgBox "Found carriage return!"
      DoEvents
    End If

    If InStr(1, strCorrect, " Asclepias sp.", vbTextCompare) > 0 Then
      DoEvents
    End If

    If InStr(1, strCorrect, "formossisimus", vbTextCompare) > 0 Then
      DoEvents
      strCorrect = Replace(strCorrect, "Erigeron formossisimus", "Erigeron formosissimus")
    End If

    If Left(strCorrect, 1) = " " Then
      strTrimName = Trim(strCorrect)
      If strCorrect <> " " Then
        DoEvents
      End If
      strCorrect = strTrimName
      strHexify = HexifyName(strTrimName)
      If MyGeneralOperations.CheckCollectionForKey(pCheckCollection, strHexify) Then
        strCorrect = pCheckCollection.Item(strHexify)
      End If
      strCorrect = Replace(strCorrect, "gramminoid", "graminoid")
      If InStr(1, strCorrect, vbCrLf) > 0 Or InStr(1, strCorrect, vbNewLine) > 0 Then
        MsgBox "Found carriage return!"
        DoEvents
      End If
    End If

    If Left(strCorrect, 1) = " " Or Left(strName, 1) = " " Or InStr(1, strCorrect, vbTab) > 0 Or InStr(1, strName, vbTab) > 0 Then
      If strName <> " " Then
        DoEvents
      End If
    End If

    If strOrigName <> strCorrect Then
      booMadeEdits = True
      strOID = CStr(pFeature.OID)
      strOID = String(4 - Len(strOID), " ") & strOID
      strEditReport = strEditReport & "  --> Feature OID " & strOID & "] Changed '" & _
          strName & "' to '" & strCorrect & "'" & vbCrLf
      strExcelReport = strExcelReport & strBase & vbTab & """" & CStr(pFeature.OID) & """" & vbTab & _
            """" & strName & """" & vbTab & """" & strCorrect & """" & vbCrLf
      pFeature.Value(lngNameIndex) = strCorrect
      pFCursor.UpdateFeature pFeature
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing

End Sub

Public Sub ReplaceNamesInShapefile(pDataset As IDataset, pCheckCollection As Collection, booMadeEdits As Boolean, _
    strEditReport As String, strBase As String, strExcelReport As String, varSpecialConversions() As Variant, _
    varQueryConversions() As Variant)

  On Error GoTo ErrHandler

  Dim pFClass As IFeatureClass
  Set pFClass = pDataset
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngNameIndex As Long
  Dim lngIndex As Long
  Dim strName As String
  Dim strReturn() As String
  Dim strHexify As String
  Dim strCorrect As String
  Dim strTrimName As String
  booMadeEdits = False
  Dim strOID As String
  Dim strOrigName As String
  Dim booIsCover As Boolean
  booIsCover = StrComp(Right(pDataset.BrowseName, 2), "_C", vbTextCompare) = 0

  lngIndex = -1

  Dim pDoneColl As New Collection

  strEditReport = "Edits to '" & pDataset.BrowseName & "':" & vbCrLf
  strExcelReport = ""

  lngNameIndex = pFClass.FindField("Species")
  If lngNameIndex = -1 Then lngNameIndex = pFClass.FindField("Cover")
  If lngNameIndex = -1 Then
    MsgBox "Unexpected Event!"
    DoEvents
    GoTo ClearMemory
  End If

  Dim strQuadrat As String
  Dim strYear As String
  Dim strCorD As String
  Dim strNoteOnChanges As String

  Call FillQuadratAndYearFromDatasetBrowsename(pDataset.BrowseName, strQuadrat, strYear, strCorD)

  If strYear = "2022" And strQuadrat = "Q8" Then
    DoEvents
  End If

  Dim pQueryFilt As IQueryFilter
  Dim strQueryString As String
  Dim varStrings() As Variant
  Dim strInstructions As String
  Dim varFeaturesToDelete() As Variant
  Dim lngDeleteIndex As Long
  Dim booYearQuadrat As Boolean
  Dim pTable As ITable

  Dim lngSpecialIndex As Long

  Set pQueryFilt = New QueryFilter

  For lngSpecialIndex = 0 To UBound(varQueryConversions, 2)

    strQueryString = ReturnQueryStringFromSpecialConversions(strQuadrat, CLng(strYear), booIsCover, _
        varQueryConversions, strInstructions, lngSpecialIndex, booYearQuadrat)

    If booYearQuadrat Then

      pQueryFilt.WhereClause = strQueryString

      If strInstructions = "Delete" Then
        DeleteFeaturesInFClassBasedOnQueryFilter pFClass, pQueryFilt, strEditReport, strExcelReport, _
              booMadeEdits, lngNameIndex, strBase

      ElseIf strInstructions = "Change Species" Then   ' THEN CHANGE SPECIES NAMES FOR SELECTED ROWS

        UpdateSpeciesInFClassBasedOnQueryFilter pFClass, pQueryFilt, varSpecialConversions, strQuadrat, _
              strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
              strBase, CStr(varQueryConversions(2, lngSpecialIndex)), CStr(varQueryConversions(3, lngSpecialIndex))

      ElseIf strInstructions = "Copy Features" Then
        CopyFeaturesInFClassBasedOnQueryFilter pFClass, strQueryString, strEditReport, strExcelReport, _
              booMadeEdits, lngNameIndex, strBase

      ElseIf strInstructions = "Return Centroid" Then
        UpdateGeometryInFClassBasedOnQueryFilter pFClass, pQueryFilt, varSpecialConversions, strQuadrat, _
              strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
              strBase, CStr(varQueryConversions(2, lngSpecialIndex)), CStr(varQueryConversions(3, lngSpecialIndex))

      End If
    End If
  Next lngSpecialIndex

  UpdateSpeciesInFClassBasedConversionArray pFClass, varSpecialConversions, strQuadrat, _
        strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
        strBase

  If SpecialConversionExistsForYearQuadrat(varSpecialConversions, strQuadrat, CLng(strYear)) Then
    UpdateSpeciesInFClassBasedConversionArray pFClass, varSpecialConversions, strQuadrat, _
          strYear, strEditReport, strExcelReport, booMadeEdits, lngNameIndex, pCheckCollection, _
          strBase
  End If

  GoTo ClearMemory
  Exit Sub

ErrHandler:
  DoEvents

ClearMemory:
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strReturn
  Set pDoneColl = Nothing
  Set pQueryFilt = Nothing
  Erase varStrings
  Erase varFeaturesToDelete
  Set pTable = Nothing

End Sub

Public Function CreateVarSpecialConversions_SA(varQueryConversions() As Variant) As Variant()

  Dim varSpecialConversions() As Variant
  Dim lngMaxIndex As Long
  Dim strConstruct As String

  lngMaxIndex = 0
  ReDim Preserve varSpecialConversions(6, lngMaxIndex)
  varSpecialConversions(0, lngMaxIndex) = "Q1000"   ' THIS CONVERSION JUST PLACEHOLDER UNTIL WE FIND REAL CHANGES
  varSpecialConversions(1, lngMaxIndex) = -999
  varSpecialConversions(2, lngMaxIndex) = "Elymus elymoides"
  varSpecialConversions(3, lngMaxIndex) = "Muhlenbergia tricholepis"
  varSpecialConversions(4, lngMaxIndex) = Array("", "")
  varSpecialConversions(5, lngMaxIndex) = "Email Margaret July 25, 2018"
  varSpecialConversions(6, lngMaxIndex) = Array("", "")

  lngMaxIndex = -1

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q24"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Gutierrezia sarothrae"
  varQueryConversions(3, lngMaxIndex) = "Gutierrezia sarothrae"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Gutierrezia sarothrae'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q22"
  varQueryConversions(1, lngMaxIndex) = 2017
  varQueryConversions(2, lngMaxIndex) = "Echinocereus engelmannii"
  varQueryConversions(3, lngMaxIndex) = "Echinocereus engelmannii"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Echinocereus engelmannii'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q21"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Menodora scabra"
  varQueryConversions(3, lngMaxIndex) = "Menodora scabra"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Menodora scabra'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q14"
  varQueryConversions(1, lngMaxIndex) = 2018
  varQueryConversions(2, lngMaxIndex) = "Bouteloua curtipendula"
  varQueryConversions(3, lngMaxIndex) = "Aristida purpurea"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" IN (28) AND ""species"" = 'Bouteloua curtipendula'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q7"
  varQueryConversions(1, lngMaxIndex) = 2020
  varQueryConversions(2, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua curtipendula"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" IN (12) AND ""species"" = 'Bouteloua hirsuta'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q4"
  varQueryConversions(1, lngMaxIndex) = 2021
  varQueryConversions(2, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(4, lngMaxIndex) = Array("Natural_Drainages_Watershed_D_Q4_2022_C|""species"" = 'Bouteloua hirsuta'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Copy Features", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q2"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Menodora scabra"
  varQueryConversions(3, lngMaxIndex) = "Menodora scabra"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Menodora scabra'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q1"
  varQueryConversions(1, lngMaxIndex) = 2021
  varQueryConversions(2, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(4, lngMaxIndex) = Array("Natural_Drainages_Watershed_A_Q1_2022_C|""species"" = 'Bouteloua hirsuta'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Copy Features", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q1"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua curtipendula"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" IN (4,5,8,9,10,11) AND ""species"" = 'Bouteloua hirsuta'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q1"
  varQueryConversions(1, lngMaxIndex) = 2020
  varQueryConversions(2, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua curtipendula"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" IN (2,3,4,5) AND ""species"" = 'Bouteloua hirsuta'", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Sierra_Ancha_1934_2024_PDF v3_Review_MMM_Dec20.xlsx"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q6"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = " "
  varQueryConversions(3, lngMaxIndex) = "Schizachyrium cirratum"
  varQueryConversions(4, lngMaxIndex) = Array("""species"" = ' '", """FID"" = -999")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q12"
  varQueryConversions(1, lngMaxIndex) = 1952
  varQueryConversions(2, lngMaxIndex) = " "
  varQueryConversions(3, lngMaxIndex) = "Townsendia exscapa"
  varQueryConversions(4, lngMaxIndex) = Array("""FID"" = -999", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q1"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q11"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q13"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q14"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q15"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q16"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q17"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q18"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q19"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q20"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Rock"
  varQueryConversions(3, lngMaxIndex) = "Deleted Rock"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Rock'", """species"" = ' '")
  varQueryConversions(5, lngMaxIndex) = "Initial Corrections; August, 2021"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q9"
  varQueryConversions(1, lngMaxIndex) = 2024
  varQueryConversions(2, lngMaxIndex) = "Abutilon palmeri"
  varQueryConversions(3, lngMaxIndex) = "Abutilon parvulum"
  varQueryConversions(4, lngMaxIndex) = Array("", """species"" = 'Abutilon palmeri'")
  varQueryConversions(5, lngMaxIndex) = "From Margaret Moore, Nov. 5 2024"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q15"
  varQueryConversions(1, lngMaxIndex) = 2024
  varQueryConversions(2, lngMaxIndex) = "Ephedra trifurca"
  varQueryConversions(3, lngMaxIndex) = "Ephedra viridis"
  varQueryConversions(4, lngMaxIndex) = Array("", """species"" = 'Ephedra trifurca'")
  varQueryConversions(5, lngMaxIndex) = "From Margaret Moore, Nov. 5 2024"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q18"
  varQueryConversions(1, lngMaxIndex) = 2024
  varQueryConversions(2, lngMaxIndex) = "Ephedra trifurca"
  varQueryConversions(3, lngMaxIndex) = "Ephedra viridis"
  varQueryConversions(4, lngMaxIndex) = Array("", """species"" = 'Ephedra trifurca'")
  varQueryConversions(5, lngMaxIndex) = "From Margaret Moore, Nov. 5 2024"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q5"
  varQueryConversions(1, lngMaxIndex) = 2024
  varQueryConversions(2, lngMaxIndex) = "Helio sp."
  varQueryConversions(3, lngMaxIndex) = "Heliomeris longifolia"
  varQueryConversions(4, lngMaxIndex) = Array("", """species"" = 'Helio sp.'")
  varQueryConversions(5, lngMaxIndex) = "From Margaret Moore, Nov. 5 2024"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q5"
  varQueryConversions(1, lngMaxIndex) = 2023
  varQueryConversions(2, lngMaxIndex) = "Aristida ternipes"
  varQueryConversions(3, lngMaxIndex) = "Aristida adscensionis"
  varQueryConversions(4, lngMaxIndex) = Array("", """species"" = 'Aristida ternipes'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Nov. 23 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q2"
  varQueryConversions(1, lngMaxIndex) = 2018
  varQueryConversions(2, lngMaxIndex) = "Bouteloua aristidoides"
  varQueryConversions(3, lngMaxIndex) = "Bouteloua hirsuta"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Bouteloua aristidoides'", """species"" = 'Bouteloua aristidoides'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q5"
  varQueryConversions(1, lngMaxIndex) = 2017
  varQueryConversions(2, lngMaxIndex) = "Eriogonum sp."
  varQueryConversions(3, lngMaxIndex) = "Eriogonum abertianum"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = ''", """species"" = 'Eriogonum sp.'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Skip", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q8"
  varQueryConversions(1, lngMaxIndex) = 2021
  varQueryConversions(2, lngMaxIndex) = "Opuntia engelmannii"
  varQueryConversions(3, lngMaxIndex) = "Opuntia phaeacantha"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Opuntia engelmannii'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q8"
  varQueryConversions(1, lngMaxIndex) = 2021
  varQueryConversions(2, lngMaxIndex) = "Opuntia engelmannii"
  varQueryConversions(3, lngMaxIndex) = "Opuntia phaeacantha"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Opuntia phaeacantha'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Return Centroid", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q8"
  varQueryConversions(1, lngMaxIndex) = 2022
  varQueryConversions(2, lngMaxIndex) = "Opuntia engelmannii"
  varQueryConversions(3, lngMaxIndex) = "Opuntia phaeacantha"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Opuntia engelmannii'", """species"" = 'Opuntia engelmannii'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q12"
  varQueryConversions(1, lngMaxIndex) = 2017
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = "Arctostaphylos sp."
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q12"
  varQueryConversions(1, lngMaxIndex) = 2018
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = "Arctostaphylos sp."
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q12"
  varQueryConversions(1, lngMaxIndex) = 2019
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = "Arctostaphylos sp."
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q12"
  varQueryConversions(1, lngMaxIndex) = 2020
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = "Arctostaphylos sp."
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q21"
  varQueryConversions(1, lngMaxIndex) = 1952
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = ""
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q21"
  varQueryConversions(1, lngMaxIndex) = 1955
  varQueryConversions(2, lngMaxIndex) = "Arctostaphylos pungens"
  varQueryConversions(3, lngMaxIndex) = ""
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Arctostaphylos pungens'", """species"" = ''")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Delete", "Skip")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q22"
  varQueryConversions(1, lngMaxIndex) = 1941
  varQueryConversions(2, lngMaxIndex) = "Bouteloua aristidoides"
  varQueryConversions(3, lngMaxIndex) = "Bothriochloa barbinodis"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Bouteloua aristidoides'", """species"" = 'Bouteloua aristidoides'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  lngMaxIndex = lngMaxIndex + 1
  ReDim Preserve varQueryConversions(6, lngMaxIndex)
  varQueryConversions(0, lngMaxIndex) = "Q24"
  varQueryConversions(1, lngMaxIndex) = 1935
  varQueryConversions(2, lngMaxIndex) = "Escobaria vivipara"
  varQueryConversions(3, lngMaxIndex) = "Opuntia engelmannii"
  varQueryConversions(4, lngMaxIndex) = Array("""Species"" = 'Escobaria vivipara'", """species"" = 'Escobaria vivipara'")
  varQueryConversions(5, lngMaxIndex) = "From Wade Gibson, Apr. 19 2023"
  varQueryConversions(6, lngMaxIndex) = Array("Change Species", "Change Species")

  CreateVarSpecialConversions_SA = varSpecialConversions
  Erase varSpecialConversions

End Function

Public Sub ReviseShapefiles_SA()

  Dim booRestrictToSite As Boolean
  Dim strSiteToRestrict As String
  booRestrictToSite = True
  strSiteToRestrict = "Natural Drainages"

  Dim varSpecialCases() As Variant
  varSpecialCases = Array("Natural_Drainages_Watershed_D_Q12_1935_D", "Natural_Drainages_Watershed_A_Q13_1940_D", _
      "Natural_Drainages_Watershed_A_Q5_2022_D", "Natural_Drainages_Watershed_D_Q4_1955_C")

  Dim pSpecialCases As New Collection
  Dim lngSpecialIndex As Long
  For lngSpecialIndex = 0 To UBound(varSpecialCases)
    pSpecialCases.Add True, CStr(varSpecialCases(lngSpecialIndex))
  Next lngSpecialIndex

  Dim strQuadrats() As String
  Dim pPlotToQuadratConversion As Collection
  Dim pQuadratToPlotConversion As Collection
  Dim lngFeatCount As Long
  Dim pQuadData As Collection
  Dim varSites() As Variant
  Dim varSitesSpecifics() As Variant
  Set pQuadData = Margaret_Functions.FillQuadratNameColl_Rev_SA(strQuadrats, pPlotToQuadratConversion, pQuadratToPlotConversion, _
      varSites, varSitesSpecifics, booRestrictToSite, strSiteToRestrict)

  Dim strItems() As String
  Dim strNote As String

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim lngMaxIndex As Long

  Dim varSpecialConversions() As Variant
  Dim varQueryConversions() As Variant
  ReDim varSpecialConversions(6, 9)
  Dim strNoteOnChanges As String

  varSpecialConversions = CreateVarSpecialConversions_SA(varQueryConversions)

  Dim pCoverCollection As New Collection
  Dim pDensityCollection As New Collection

  Dim pCoverToDensity As Collection
  Dim pDensityToCover As Collection
  Dim strCoverToDensityQuery As String
  Dim strDensityToCoverQuery As String
  Dim pCoverShouldChangeColl As Collection
  Dim pDensityShouldChangeColl As Collection

  Debug.Print "---------------------"
  Call FillCollections_SA(pCoverCollection, pDensityCollection, pCoverToDensity, pDensityToCover, _
    strCoverToDensityQuery, strDensityToCoverQuery, pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim pQueryFilt As IQueryFilter
  Dim strQueryString As String
  Dim varStrings() As Variant

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray
  Dim strRoot As String
  Dim strContainingFolder As String
  Call DeclareWorkspaces(strRoot, , , , , strContainingFolder)

  Set pFolders = MyGeneralOperations.ReturnFoldersFromNestedFolders(strRoot, "")
  Dim strFolder As String
  Dim lngIndex As Long

  For lngIndex = 0 To pFolders.Count - 1
    Debug.Print CStr(lngIndex) & "]  " & pFolders.Element(lngIndex)
  Next lngIndex

  Dim pDataset As IDataset
  Dim booFoundShapefiles As Boolean
  Dim varDatasets() As Variant

  Dim strNames() As String
  Dim strName As String
  Dim lngDatasetIndex As Long
  Dim lngNameIndex As Long
  Dim lngNameCount As Long
  Dim booFoundNames As Boolean
  Dim lngRecCount As Long

  Dim strFullNames() As String
  Dim lngFullNameCounter As Long

  Dim lngShapefileCount As Long
  Dim lngAcceptSFCount As Long
  lngShapefileCount = 0
  lngRecCount = 0
  lngAcceptSFCount = 0

  lngFullNameCounter = -1
  Dim pNameColl As New Collection
  Dim strHexify As String
  Dim strCorrect As String
  Dim pCheckCollection As Collection
  Dim strReport As String
  Dim booMadeChanges As Boolean
  Dim strEditReport As String
  Dim strExcelReport As String
  Dim strExcelFullReport As String
  Dim pFClass As IFeatureClass
  Dim strBase As String
  Dim strSplit() As String
  Dim strQuadrat As String
  Dim strYear As String
  Dim strCorD As String

  Dim pGeoDataset As IGeoDataset

  strExcelFullReport = """Shapefile""" & vbTab & """Quadrat""" & vbTab & """Year""" & vbTab & _
      """Type""" & vbTab & """Feature_ID""" & vbTab & """Original""" & vbTab & """Changed_To""" & vbCrLf

  Dim strExcelResizeReport As String
  strExcelResizeReport = ""

  For lngIndex = 0 To pFolders.Count - 1
    DoEvents
    strFolder = pFolders.Element(lngIndex)
    varDatasets = ReturnFeatureClassesOrNothing(strFolder, booFoundShapefiles)

    Debug.Print CStr(lngIndex + 1) & " of " & CStr(pFolders.Count) & "] " & strFolder
    If booFoundShapefiles Then
      Debug.Print "  --> Found Shapefiles = " & CStr(booFoundShapefiles) & " [n = " & CStr(UBound(varDatasets) + 1) & "]"

      lngShapefileCount = lngShapefileCount + UBound(varDatasets) + 1

      For lngDatasetIndex = 0 To UBound(varDatasets)
        Set pDataset = varDatasets(lngDatasetIndex)
        Set pGeoDataset = pDataset

        If Not pGeoDataset.Extent.IsEmpty Then
          If pGeoDataset.Extent.XMax > 1.5 Or pGeoDataset.Extent.YMax > 1.5 Then
            Debug.Print "Extent too Large:  " & pDataset.BrowseName & "...Max X = " & Format(pGeoDataset.Extent.XMax, "0.000") & ", " & _
                "Max Y = " & Format(pGeoDataset.Extent.YMax, "0.000")
            strExcelResizeReport = strExcelResizeReport & pDataset.BrowseName & "...Max X = " & Format(pGeoDataset.Extent.XMax, "0.000") & ", " & _
                "Max Y = " & Format(pGeoDataset.Extent.YMax, "0.000") & vbCrLf
          End If
        End If

        If MyGeneralOperations.CheckCollectionForKey(pSpecialCases, pDataset.BrowseName) Then
          TransformFeaturesFrom_10x10_to_1x1 pDataset
        End If

        If Right(pDataset.BrowseName, 2) = "_D" Then
          Set pCheckCollection = pDensityCollection
        ElseIf Right(pDataset.BrowseName, 2) = "_C" Then
          Set pCheckCollection = pCoverCollection
        Else
          MsgBox "Unexpected Dataset Name!"
          DoEvents
        End If
        Call FillQuadratAndYearFromDatasetBrowsename(pDataset.BrowseName, strQuadrat, strYear, strCorD)

        If strQuadrat = "Q2" And strYear = "2018" Then  'And strCorD = "C" Then
          DoEvents
        End If

        strBase = """" & pDataset.BrowseName & """" & vbTab & """" & strQuadrat & """" & vbTab & _
            """" & strYear & """" & vbTab & """" & IIf(strCorD = "C", "Cover", "Density") & """"

        Set pFClass = pDataset
        If pFClass.FindField("Cover") > -1 Or pFClass.FindField("Species") > -1 Then

          Call ReplaceNamesInShapefile(pDataset, pCheckCollection, booMadeChanges, strEditReport, strBase, _
              strExcelReport, varSpecialConversions, varQueryConversions)

          If booMadeChanges Then
            strReport = strReport & strEditReport
            strExcelFullReport = strExcelFullReport & strExcelReport
          Else
            strReport = strReport & "No changes to '" & pDataset.BrowseName & "'..." & vbCrLf
            strExcelFullReport = strExcelFullReport & strBase & vbTab & """<- No Changes ->""" & vbTab & vbTab & vbCrLf
          End If

        End If
      Next lngDatasetIndex

    End If

  Next lngIndex

  strReport = strReport & vbCrLf & "Done..." & vbCrLf & _
    MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

  Dim pDataObj As New MSForms.DataObject

  strExcelFullReport = Replace(strExcelFullReport, vbTab, ",")
  MyGeneralOperations.WriteTextFile strContainingFolder & "\Log_of_Changes_" & MyGeneralOperations.ReturnTimeStamp & ".csv", strExcelFullReport

  MyGeneralOperations.WriteTextFile strContainingFolder & "\Extents_Too_Large_" & MyGeneralOperations.ReturnTimeStamp & ".csv", strExcelResizeReport

  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Erase strQuadrats
  Set pPlotToQuadratConversion = Nothing
  Set pQuadratToPlotConversion = Nothing
  Set pQuadData = Nothing
  Erase strItems
  Erase varSpecialConversions
  Erase varQueryConversions
  Set pCoverCollection = Nothing
  Set pDensityCollection = Nothing
  Set pCoverToDensity = Nothing
  Set pDensityToCover = Nothing
  Set pCoverShouldChangeColl = Nothing
  Set pDensityShouldChangeColl = Nothing
  Set pQueryFilt = Nothing
  Erase varStrings
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Erase varDatasets
  Erase strNames
  Erase strFullNames
  Set pNameColl = Nothing
  Set pCheckCollection = Nothing
  Set pFClass = Nothing
  Erase strSplit
  Set pDataObj = Nothing

End Sub

Public Sub TransformFeaturesFrom_10x10_to_1x1(pFClass As IFeatureClass)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pGeom As IGeometry
  Dim pPtColl As IPointCollection
  Dim pPoint As IPoint
  Dim lngIndex As Long
  Dim pFClassManage As IFeatureClassManage
  Dim pSrcPoint As IPoint
  Dim pSrcPolygon As IPolygon

  Dim pMaxExtent As IEnvelope
  Dim pGeoDataset As IGeoDataset
  Set pGeoDataset = pFClass
  Set pMaxExtent = pGeoDataset.Extent
  If pMaxExtent.XMax > 1.1 Or pMaxExtent.YMax > 1.1 Then

    Set pPoint = New Point
    Set pFCursor = pFClass.Update(Nothing, False)
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      Set pGeom = pFeature.ShapeCopy
      If TypeOf pGeom Is IPoint Then
        Set pSrcPoint = pGeom
        pSrcPoint.x = pSrcPoint.x / 10
        pSrcPoint.Y = pSrcPoint.Y / 10

        Set pFeature.Shape = pSrcPoint

      ElseIf TypeOf pGeom Is IPolygon Then
        Set pSrcPolygon = pGeom
        Set pPtColl = pSrcPolygon

        For lngIndex = 0 To pPtColl.PointCount - 1
          pPtColl.QueryPoint lngIndex, pPoint
          pPoint.x = pPoint.x / 10
          pPoint.Y = pPoint.Y / 10
          pPtColl.UpdatePoint lngIndex, pPoint
        Next lngIndex

        Set pFeature.Shape = pGeom
      End If

      pFCursor.UpdateFeature pFeature
      pFCursor.Flush

      Set pFeature = pFCursor.NextFeature
    Loop

    Set pFClassManage = pFClass
    pFClassManage.UpdateExtent

  End If

ClearMemory:
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pGeom = Nothing
  Set pPtColl = Nothing
  Set pPoint = Nothing
  Set pFClassManage = Nothing

End Sub

Public Sub AddVerbatimFields_SA(pFClass As IFeatureClass, pQuadData As Collection)

  Dim strName As String
  Dim pFCursor As IFeatureCursor
  Dim lngSrcSpeciesNameIndex As Long
  Dim lngVerbSpeciesNameIndex As Long
  Dim lngRotationNameIndex As Long
  Dim lngVerbTypeIndex As Long
  Dim lngSiteIndex As Long
  Dim lngPlotIndex As Long

  Dim pField As iField
  Dim pFieldEdit As IFieldEdit
  Dim pFeature As IFeature
  Dim pDataset As IDataset

  Dim strQuad As String
  Dim varItems() As Variant
  Dim strSite As String
  Dim strPlot As String
  Dim strFileHeader As String
  Set pDataset = pFClass

  strQuad = aml_func_mod.ReturnFilename2(pDataset.Workspace.PathName)
  strQuad = Right(strQuad, Len(strQuad) - InStr(1, strQuad, "Q"))
  strQuad = Replace(strQuad, "Q", "", , , vbTextCompare)
  varItems = pQuadData.Item(strQuad)
  strSite = Trim(varItems(1))
  If strSite = "" Then
    strSite = Trim(varItems(0))
  End If
  strPlot = Trim(varItems(2))
  strFileHeader = Trim(varItems(5))

  lngSrcSpeciesNameIndex = pFClass.FindField("Species")
  lngVerbSpeciesNameIndex = pFClass.FindField("Verb_Spcs")
  lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  lngSiteIndex = pFClass.FindField("Site")
  lngPlotIndex = pFClass.FindField("Plot")

  If lngSiteIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Site"
      .Type = esriFieldTypeString
      .length = 75
    End With
    pFClass.AddField pField
    lngSiteIndex = pFClass.FindField("Site")
  End If
  If lngPlotIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Plot"
      .Type = esriFieldTypeString
      .length = 25
    End With
    pFClass.AddField pField
    lngPlotIndex = pFClass.FindField("Plot")
  End If
  If lngVerbSpeciesNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Verb_Spcs"
      .Type = esriFieldTypeString
      .length = 50
    End With
    pFClass.AddField pField
    lngVerbSpeciesNameIndex = pFClass.FindField("Verb_Spcs")
  End If
  lngVerbTypeIndex = pFClass.FindField("Verb_Type")
  If lngVerbTypeIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Verb_Type"
      .Type = esriFieldTypeString
      .length = 50
    End With
    pFClass.AddField pField
    lngVerbTypeIndex = pFClass.FindField("Verb_Type")
  End If

  lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  If lngRotationNameIndex = -1 Then
    Set pField = New Field
    Set pFieldEdit = pField
    With pFieldEdit
      .Name = "Revise_Rtn"
      .Type = esriFieldTypeDouble
      .Precision = 12
      .Scale = 6
    End With
    pFClass.AddField pField
    lngRotationNameIndex = pFClass.FindField("Revise_Rtn")
  End If

  Set pFCursor = pFClass.Update(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pFeature.Value(lngVerbSpeciesNameIndex) = pFeature.Value(lngSrcSpeciesNameIndex)
    If StrComp(Right(pDataset.BrowseName, 2), "_C", vbTextCompare) = 0 Then
      pFeature.Value(lngVerbTypeIndex) = "Cover"
    Else
      pFeature.Value(lngVerbTypeIndex) = "Density"
    End If
    pFeature.Value(lngSiteIndex) = strSite
    pFeature.Value(lngPlotIndex) = strPlot
    pFCursor.UpdateFeature pFeature
    Set pFeature = pFCursor.NextFeature
  Loop
  pFCursor.Flush

ClearMemory:
  Set pFCursor = Nothing
  Set pField = Nothing
  Set pFieldEdit = Nothing
  Set pFeature = Nothing
  Set pDataset = Nothing

End Sub

Public Sub FillCollections_SA(pCoverCollection As Collection, pDensityCollection As Collection, _
    Optional pCoverToDensity As Collection, Optional pDensityToCover As Collection, _
    Optional strCoverToDensityQuery As String, Optional strDensityToCoverQuery As String, _
    Optional pCoverShouldChangeColl As Collection, Optional pDensityShouldChangeColl As Collection, _
    Optional pCoverCommentCollection As Collection, Optional pDensityCommentCollection As Collection)

  Dim pRedigitizeColl As Collection
  Set pRedigitizeColl = ReturnReplacementColl

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pTable As ITable
  Set pWSFact = New ExcelWorkspaceFactory

  Dim pTestWS As IFeatureWorkspace
  Dim pTestWSFact As IWorkspaceFactory
  Set pTestWSFact = New FileGDBWorkspaceFactory
  Set pTestWS = pTestWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data\Combined_by_Quadrat.gdb", 0)
  Dim pTestFC As IFeatureClass
  Set pTestFC = pTestWS.OpenFeatureClass("Box")
  Dim strPrefix As String
  Dim strSuffix As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pTestFC, strPrefix, strSuffix)

  Dim lngCorrectIndex As Long
  Dim lngIncorrectIndex As Long
  Dim strCorrect As String
  Dim strIncorrect As String
  Dim strHexCorrect As String
  Dim strHexIncorrect As String
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngShouldChangeIndex As Long
  Dim strShouldChange As String
  Dim booShouldChange As Boolean

  Dim varWorksheets As Variant
  Dim varColls As Variant
  Dim varVals As Variant
  Dim varShouldChange As Variant
  Dim strCover() As String
  Dim strDensity() As String

  Dim strPath As String
  strPath = "D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Summary_Data_from_JSJ\Sierra_Ancha_Species_Lists_Aug 2021_MMM_jsj.xlsx"

  varWorksheets = Array("Cover_Species$", "Density_Species$")
  varColls = Array(pCoverCollection, pDensityCollection)
  varVals = Array(strCover, strDensity)
  Set pCoverShouldChangeColl = New Collection
  Set pDensityShouldChangeColl = New Collection
  varShouldChange = Array(pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim lngIndex As Long
  Dim pColl As Collection
  Dim strVals() As String
  Dim varVal As Variant
  Dim strIncorrectVariant As String

  Dim pFromCoverColl As New Collection
  Dim strFromCoverVals() As String
  Dim pFromDensColl As New Collection
  Dim pShouldChangeColl As New Collection
  Dim strFromDensVals() As String
  Dim lngFromCoverCounter As Long
  Dim lngFromDensCounter As Long
  Dim booShouldChangeFromCover As Boolean
  Dim booShouldChangeFromDensity As Boolean
  Dim lngCommentIndex As Long
  Dim strComment As String

  Set pCoverToDensity = New Collection
  Set pDensityToCover = New Collection
  strCoverToDensityQuery = ""
  strDensityToCoverQuery = ""
  lngFromCoverCounter = -1
  lngFromDensCounter = -1

  Dim strWorksheet As String

  Dim lngIndex2 As Long
  Set pWS = pWSFact.OpenFromFile(strPath, 0)

  For lngIndex = 0 To UBound(varWorksheets)
    strWorksheet = varWorksheets(lngIndex)
    Set pColl = varColls(lngIndex)
    strVals = varVals(lngIndex)
    Set pShouldChangeColl = varShouldChange(lngIndex)
    lngIndex2 = -1

    Set pTable = pWS.OpenTable(strWorksheet)
    lngCorrectIndex = pTable.FindField("Correct")
    lngIncorrectIndex = pTable.FindField("Incorrect")
    lngShouldChangeIndex = pTable.FindField("Comment")
    lngCommentIndex = pTable.FindField("Comment_2")
    If lngShouldChangeIndex = -1 Then lngShouldChangeIndex = pTable.FindField("Comments1")
    Debug.Print CStr(lngIndex) & "] " & IIf(lngIndex = 0, "Cover", "Density") & " Record Count = " & CStr(pTable.RowCount(Nothing))
    Set pCursor = pTable.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      strCorrect = ""
      strIncorrect = ""
      strShouldChange = ""
      booShouldChangeFromCover = False
      booShouldChangeFromDensity = False
      booShouldChange = False

      varVal = pRow.Value(lngCorrectIndex)
      If Not IsNull(varVal) Then strCorrect = CStr(varVal)
      varVal = pRow.Value(lngIncorrectIndex)
      If Not IsNull(varVal) Then strIncorrect = CStr(varVal)
      varVal = pRow.Value(lngShouldChangeIndex)
      If Not IsNull(varVal) Then strShouldChange = Trim(CStr(varVal))

      If strIncorrect = "Lycurus phleoides" Or strCorrect = "Ambrosia Sp." Then
        DoEvents
      End If

      If InStr(1, strShouldChange, "change to", vbTextCompare) > 0 Then booShouldChange = True

      strComment = MyGeneralOperations.ReturnStringValFromRow(pRow, lngCommentIndex)

      Set pRow = pCursor.NextRow

      If strShouldChange <> "" Then

        If strShouldChange = "change to point shapefile" Then
          booShouldChangeFromCover = True
        ElseIf strShouldChange = "change to polygon shapefile" Then
          booShouldChangeFromDensity = True
        End If

        If lngIndex = 0 Then  ' cover
          If Not MyGeneralOperations.CheckCollectionForKey(pFromCoverColl, strShouldChange) Then
            lngFromCoverCounter = lngFromCoverCounter + 1
            ReDim Preserve strFromCoverVals(lngFromCoverCounter)
            strFromCoverVals(lngFromCoverCounter) = strShouldChange
            pFromCoverColl.Add True, strShouldChange
          End If
        ElseIf lngIndex = 1 Then  ' density
          If Not MyGeneralOperations.CheckCollectionForKey(pFromDensColl, strShouldChange) Then
            lngFromDensCounter = lngFromDensCounter + 1
            ReDim Preserve strFromDensVals(lngFromDensCounter)
            strFromDensVals(lngFromDensCounter) = strShouldChange
            pFromDensColl.Add True, strShouldChange
          End If
        End If
      End If

      If InStr(1, strIncorrect, "Erigeron formosissimus", vbTextCompare) > 0 Then
        DoEvents
      End If

      strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
      strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
      strCorrect = Trim(strCorrect)

      If InStr(1, strCorrect, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
        DoEvents
      End If

      If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strCorrect = " & strCorrect
      End If
      If InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strIncorrect = " & strIncorrect
      End If

      strHexCorrect = HexifyName(strCorrect)
      strHexIncorrect = HexifyName(strIncorrect)
      If Not pCoverCommentCollection Is Nothing And Not pDensityCommentCollection Is Nothing Then
        If lngIndex = 0 Then
          pCoverCommentCollection.Add strComment, strHexIncorrect
        Else
          pDensityCommentCollection.Add strComment, strHexIncorrect
        End If
      End If

      If strCorrect = "" Then
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexIncorrect) Then
           pShouldChangeColl.Add booShouldChange, strHexIncorrect  ' strHexIncorrect is the correct name in this case
        End If

      ElseIf strCorrect <> "" Then
        If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
          pColl.Add strCorrect, strHexIncorrect
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
          pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
        End If

        lngIndex2 = lngIndex2 + 1
        ReDim Preserve strVals(lngIndex2)
        strVals(lngIndex2) = strIncorrect

        If lngIndex = 0 And booShouldChangeFromCover Then
          If Not MyGeneralOperations.CheckCollectionForKey(pCoverToDensity, strHexIncorrect) Then
            pCoverToDensity.Add strCorrect, strHexIncorrect
          End If

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          If Not MyGeneralOperations.CheckCollectionForKey(pDensityToCover, strHexIncorrect) Then
            pDensityToCover.Add strCorrect, strHexIncorrect
          End If
          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        End If

        If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Then
          strIncorrectVariant = strIncorrect

          strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
          strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbCrLf)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbNewLine)), "")

          strHexCorrect = HexifyName(strCorrect)
          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If

          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

        If Trim(strIncorrect) <> strIncorrect Then
          strIncorrectVariant = Trim(strIncorrect)

          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If
          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

      Else

        If lngIndex = 0 And booShouldChangeFromCover Then
          pCoverToDensity.Add strIncorrect, strIncorrect

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          pDensityToCover.Add strIncorrect, strIncorrect

          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        End If
      End If
    Loop
    varVals(lngIndex) = strVals
  Next lngIndex

  Debug.Print "Checking From Cover to Density Values:"
  For lngIndex = 0 To lngFromCoverCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromCoverVals(lngIndex)
  Next lngIndex
  Debug.Print "Checking From Density to Cover Values:"
  For lngIndex = 0 To lngFromDensCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromDensVals(lngIndex)
  Next lngIndex

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varWorksheets = Null
  varColls = Null
  Set pColl = Nothing

End Sub

Public Sub FillCollections(pCoverCollection As Collection, pDensityCollection As Collection, _
    Optional pCoverToDensity As Collection, Optional pDensityToCover As Collection, _
    Optional strCoverToDensityQuery As String, Optional strDensityToCoverQuery As String, _
    Optional pCoverShouldChangeColl As Collection, Optional pDensityShouldChangeColl As Collection)

  Dim pRedigitizeColl As Collection
  Set pRedigitizeColl = ReturnReplacementColl

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pTable As ITable
  Set pWSFact = New ExcelWorkspaceFactory

  Dim pTestWS As IFeatureWorkspace
  Dim pTestWSFact As IWorkspaceFactory
  Set pTestWSFact = New FileGDBWorkspaceFactory
  Set pTestWS = pTestWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data\Combined_by_Quadrat.gdb", 0)
  Dim pTestFC As IFeatureClass
  Set pTestFC = pTestWS.OpenFeatureClass("Box")
  Dim strPrefix As String
  Dim strSuffix As String
  Call MyGeneralOperations.ReturnQuerySpecialCharacters(pTestFC, strPrefix, strSuffix)

  Dim lngCorrectIndex As Long
  Dim lngIncorrectIndex As Long
  Dim strCorrect As String
  Dim strIncorrect As String
  Dim strHexCorrect As String
  Dim strHexIncorrect As String
  Dim pCursor As ICursor
  Dim pRow As IRow
  Dim lngShouldChangeIndex As Long
  Dim strShouldChange As String
  Dim booShouldChange As Boolean

  Dim varPaths As Variant
  Dim varColls As Variant
  Dim varVals As Variant
  Dim varShouldChange As Variant
  Dim strCover() As String
  Dim strDensity() As String

  varPaths = Array("D:\arcGIS_stuff\consultation\Margaret_Moore\species_list_Cover_changes_Dec_2_2017.xlsx", _
                   "D:\arcGIS_stuff\consultation\Margaret_Moore\Species_list_Density_changes_Dec_2_2017.xlsx")
  varColls = Array(pCoverCollection, pDensityCollection)
  varVals = Array(strCover, strDensity)
  Set pCoverShouldChangeColl = New Collection
  Set pDensityShouldChangeColl = New Collection
  varShouldChange = Array(pCoverShouldChangeColl, pDensityShouldChangeColl)

  Dim lngIndex As Long
  Dim strPath As String
  Dim pColl As Collection
  Dim strVals() As String
  Dim varVal As Variant
  Dim strIncorrectVariant As String

  Dim pFromCoverColl As New Collection
  Dim strFromCoverVals() As String
  Dim pFromDensColl As New Collection
  Dim pShouldChangeColl As New Collection
  Dim strFromDensVals() As String
  Dim lngFromCoverCounter As Long
  Dim lngFromDensCounter As Long
  Dim booShouldChangeFromCover As Boolean
  Dim booShouldChangeFromDensity As Boolean

  Set pCoverToDensity = New Collection
  Set pDensityToCover = New Collection
  strCoverToDensityQuery = ""
  strDensityToCoverQuery = ""
  lngFromCoverCounter = -1
  lngFromDensCounter = -1

  Dim lngIndex2 As Long

  For lngIndex = 0 To UBound(varPaths)
    strPath = varPaths(lngIndex)
    Set pColl = varColls(lngIndex)
    strVals = varVals(lngIndex)
    Set pShouldChangeColl = varShouldChange(lngIndex)
    lngIndex2 = -1

    Set pWS = pWSFact.OpenFromFile(strPath, 0)
    Set pTable = pWS.OpenTable("For_ArcGIS_Dec_2017$")
    lngCorrectIndex = pTable.FindField("Correct")
    lngIncorrectIndex = pTable.FindField("Incorrect")
    lngShouldChangeIndex = pTable.FindField("Comment")
    If lngShouldChangeIndex = -1 Then lngShouldChangeIndex = pTable.FindField("Comments1")
    Debug.Print CStr(lngIndex) & "] " & IIf(lngIndex = 0, "Cover", "Density") & " Record Count = " & CStr(pTable.RowCount(Nothing))
    Set pCursor = pTable.Search(Nothing, False)
    Set pRow = pCursor.NextRow
    Do Until pRow Is Nothing
      strCorrect = ""
      strIncorrect = ""
      strShouldChange = ""
      booShouldChangeFromCover = False
      booShouldChangeFromDensity = False
      booShouldChange = False

      varVal = pRow.Value(lngCorrectIndex)
      If Not IsNull(varVal) Then strCorrect = CStr(varVal)
      varVal = pRow.Value(lngIncorrectIndex)
      If Not IsNull(varVal) Then strIncorrect = CStr(varVal)
      varVal = pRow.Value(lngShouldChangeIndex)
      If Not IsNull(varVal) Then strShouldChange = Trim(CStr(varVal))

      If strIncorrect = "Drymaria leptophyllum" Then
        DoEvents
      End If

      If InStr(1, strShouldChange, "change to", vbTextCompare) > 0 Then booShouldChange = True

      Set pRow = pCursor.NextRow

      If strShouldChange <> "" Then

        If strShouldChange = "change to point shapefile" Then
          booShouldChangeFromCover = True
        ElseIf strShouldChange = "change to polygon shapefile" Then
          booShouldChangeFromDensity = True
        End If

        If lngIndex = 0 Then  ' cover
          If Not MyGeneralOperations.CheckCollectionForKey(pFromCoverColl, strShouldChange) Then
            lngFromCoverCounter = lngFromCoverCounter + 1
            ReDim Preserve strFromCoverVals(lngFromCoverCounter)
            strFromCoverVals(lngFromCoverCounter) = strShouldChange
            pFromCoverColl.Add True, strShouldChange
          End If
        ElseIf lngIndex = 1 Then  ' density
          If Not MyGeneralOperations.CheckCollectionForKey(pFromDensColl, strShouldChange) Then
            lngFromDensCounter = lngFromDensCounter + 1
            ReDim Preserve strFromDensVals(lngFromDensCounter)
            strFromDensVals(lngFromDensCounter) = strShouldChange
            pFromDensColl.Add True, strShouldChange
          End If
        End If
      End If

      If InStr(1, strIncorrect, "Erigeron formosissimus", vbTextCompare) > 0 Then
        DoEvents
      End If

      strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
      strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
      strCorrect = Trim(strCorrect)

      If InStr(1, strCorrect, "Muhlenbergia tricholepis", vbTextCompare) > 0 Then
        DoEvents
      End If

      If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strCorrect = " & strCorrect
      End If
      If InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbTab)), vbTextCompare) > 0 Then
        Debug.Print "...strIncorrect = " & strIncorrect
      End If

      If strCorrect = "" Then
        strHexIncorrect = HexifyName(strIncorrect)
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexIncorrect) Then
           pShouldChangeColl.Add booShouldChange, strHexIncorrect  ' strHexIncorrect is the correct name in this case
        End If

      ElseIf strCorrect <> "" Then
        strHexCorrect = HexifyName(strCorrect)
        strHexIncorrect = HexifyName(strIncorrect)
        If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
          pColl.Add strCorrect, strHexIncorrect
        End If
        If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
          pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
        End If

        lngIndex2 = lngIndex2 + 1
        ReDim Preserve strVals(lngIndex2)
        strVals(lngIndex2) = strIncorrect

        If lngIndex = 0 And booShouldChangeFromCover Then
          If Not MyGeneralOperations.CheckCollectionForKey(pCoverToDensity, strHexIncorrect) Then
            pCoverToDensity.Add strCorrect, strHexIncorrect
          End If

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          If Not MyGeneralOperations.CheckCollectionForKey(pDensityToCover, strHexIncorrect) Then
            pDensityToCover.Add strCorrect, strHexIncorrect
          End If
          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

        End If

        If InStr(1, strCorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strCorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbCrLf)), vbTextCompare) > 0 Or _
            InStr(1, strIncorrect, Chr(Asc(vbNewLine)), vbTextCompare) > 0 Then
          strIncorrectVariant = strIncorrect

          strCorrect = Replace(strCorrect, Chr(Asc(vbCrLf)), "")
          strCorrect = Replace(strCorrect, Chr(Asc(vbNewLine)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbCrLf)), "")
          strIncorrectVariant = Replace(strIncorrectVariant, Chr(Asc(vbNewLine)), "")

          strHexCorrect = HexifyName(strCorrect)
          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If

          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

        If Trim(strIncorrect) <> strIncorrect Then
          strIncorrectVariant = Trim(strIncorrect)

          strHexIncorrect = HexifyName(strIncorrectVariant)
          If Not MyGeneralOperations.CheckCollectionForKey(pColl, strHexIncorrect) Then
            pColl.Add strCorrect, strHexIncorrect

            lngIndex2 = lngIndex2 + 1
            ReDim Preserve strVals(lngIndex2)
            strVals(lngIndex2) = strIncorrectVariant

            If lngIndex = 0 And booShouldChangeFromCover Then
              pCoverToDensity.Add strCorrect, strHexIncorrect

              strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"

            ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
              pDensityToCover.Add strCorrect, strHexIncorrect

              strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                      strPrefix & "Species" & strSuffix & " = '" & strCorrect & "'"
            End If
          End If
          If Not MyGeneralOperations.CheckCollectionForKey(pShouldChangeColl, strHexCorrect) Then
            pShouldChangeColl.Add booShouldChange, strHexCorrect  ' strHexCorrect is the correct name in this case
          End If
        End If

      Else

        If lngIndex = 0 And booShouldChangeFromCover Then
          pCoverToDensity.Add strIncorrect, strIncorrect

          strCoverToDensityQuery = strCoverToDensityQuery & IIf(strCoverToDensityQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        ElseIf lngIndex = 1 And booShouldChangeFromDensity Then
          pDensityToCover.Add strIncorrect, strIncorrect

          strDensityToCoverQuery = strDensityToCoverQuery & IIf(strDensityToCoverQuery = "", "", " OR ") & _
                  strPrefix & "Species" & strSuffix & " = '" & strIncorrect & "'"

        End If
      End If
    Loop
    varVals(lngIndex) = strVals
  Next lngIndex

  Debug.Print "Checking From Cover to Density Values:"
  For lngIndex = 0 To lngFromCoverCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromCoverVals(lngIndex)
  Next lngIndex
  Debug.Print "Checking From Density to Cover Values:"
  For lngIndex = 0 To lngFromDensCounter
    Debug.Print "  " & CStr(lngIndex + 1) & "] " & strFromDensVals(lngIndex)
  Next lngIndex

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pTable = Nothing
  Set pCursor = Nothing
  Set pRow = Nothing
  varPaths = Null
  varColls = Null
  Set pColl = Nothing

End Sub

Public Function ReturnReplacementColl_SA() As Collection

  Dim pReturn As New Collection
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Dim pDataset As IDataset
  Dim pEnumDataset As IEnumDataset
  Dim strDatasetName As String

  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Sierra_Ancha\Original_Data\Redigitized_Data.gdb", 0)

  Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
  Set pDataset = pEnumDataset.Next
  Do Until pDataset Is Nothing
    strDatasetName = pDataset.BrowseName
    Select Case strDatasetName
      Case Else
        pReturn.Add pDataset, strDatasetName
    End Select
    Set pDataset = pEnumDataset.Next
  Loop

  Set ReturnReplacementColl_SA = pReturn

  Set pReturn = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDataset = Nothing
  Set pEnumDataset = Nothing

End Function

Public Function ReturnReplacementColl() As Collection

  Dim pReturn As New Collection
  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory

  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile("D:\arcGIS_stuff\consultation\Margaret_Moore\Newly_Georeferenced_Aug_2018\New_Feature_Classes.gdb", 0)

  Dim pDataset As IDataset
  Dim pEnumDataset As IEnumDataset

  Dim strDatasetName As String

  Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
  Set pDataset = pEnumDataset.Next
  Do Until pDataset Is Nothing
    strDatasetName = pDataset.BrowseName
    Select Case strDatasetName
      Case "BS_2004_46_C"
        pReturn.Add pDataset, "Q9_2015_C"
      Case "BS_2004_46_D"
        pReturn.Add pDataset, "Q9_2015_D"
      Case Else
        pReturn.Add pDataset, strDatasetName
    End Select
    Set pDataset = pEnumDataset.Next
  Loop

  Set ReturnReplacementColl = pReturn

  Set pReturn = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDataset = Nothing
  Set pEnumDataset = Nothing

End Function

Public Function HexifyName(strName As String) As String
  Dim lngIndex As Long
  HexifyName = Space$(Len(strName) * 4)
  For lngIndex = 0 To Len(strName) - 1
      Mid$(HexifyName, lngIndex * 4 + 1, 4) = Right$("0000" & Hex$(AscW(Mid$(strName, lngIndex + 1, 1))), 4)
  Next lngIndex
End Function

Public Function ReturnFeatureClassesOrNothing(strFolder As String, booWorked As Boolean, _
    Optional booFoundPolygonFClass As Boolean, Optional booFoundPointFClass As Boolean, _
    Optional pRepPointFClass As IFeatureClass, Optional pRepPolyFClass As IFeatureClass) As Variant()

  On Error GoTo ErrHandler

  Dim pWS As IWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New ShapefileWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
  Dim pEnumDataset As IEnumDataset
  Dim lngIndex As Long
  Dim varReturn() As Variant

  Dim pDataset As IDataset
  Dim pFClass As IFeatureClass

  booFoundPolygonFClass = False
  booFoundPointFClass = False

  lngIndex = -1

  Dim lngMaxPolygonFCount As Long
  Dim lngMaxPointFCount As Long
  lngMaxPolygonFCount = -1
  lngMaxPointFCount = -1

  Set pEnumDataset = pWS.Datasets(esriDTFeatureClass)
  pEnumDataset.Reset
  Set pDataset = pEnumDataset.Next
  Do Until pDataset Is Nothing
    lngIndex = lngIndex + 1
    ReDim Preserve varReturn(lngIndex)
    Set pFClass = pDataset
    If pFClass.ShapeType = esriGeometryPoint Then
      booFoundPointFClass = True
      If pFClass.Fields.FieldCount > lngMaxPointFCount Then
        Set pRepPointFClass = pFClass
        lngMaxPointFCount = pFClass.Fields.FieldCount
      End If
    ElseIf pFClass.ShapeType = esriGeometryPolygon Then
      booFoundPolygonFClass = True
      If pFClass.Fields.FieldCount > lngMaxPolygonFCount Then
        Set pRepPolyFClass = pFClass
        lngMaxPolygonFCount = pFClass.Fields.FieldCount
      End If
    End If
    Set varReturn(lngIndex) = pDataset
    Set pDataset = pEnumDataset.Next
  Loop

  ReturnFeatureClassesOrNothing = varReturn

  booWorked = lngIndex > -1
  GoTo ClearMemory
  Exit Function

ErrHandler:
  booWorked = False

ClearMemory:
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pEnumDataset = Nothing
  Erase varReturn
  Set pDataset = Nothing
  Set pFClass = Nothing

End Function

Public Function CopyFeatureClassToShapefile(pFClass As IFeatureClass, strPath As String) As IFeatureClass

  Dim strTemp As String
  strTemp = aml_func_mod.SetExtension(strPath, "shp")
  If aml_func_mod.ExistFileDir(strTemp) Then
    Debug.Print "...Feature Class '" & strPath & "' already exists.  Did not export..."
    GoTo ClearMemory
  End If

  Dim strDir As String
  Dim strFClassName As String
  strDir = aml_func_mod.ReturnDir3(strPath, False)
  strFClassName = aml_func_mod.ReturnFilename2(strPath)
  strFClassName = aml_func_mod.ClipExtension2(strFClassName)

  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Dim pFieldArray As esriSystem.IVariantArray
  Dim pField As iField
  Dim lngIndex As Long
  Dim pClone As IClone
  Dim pFieldEdit As IFieldEdit
  Dim pNewField As iField
  Dim pSourceNameColl As New Collection
  Dim strNewName As String
  Dim strSourceName As String

  Set pFieldArray = New esriSystem.varArray

  For lngIndex = 0 To pFClass.Fields.FieldCount - 1
    Set pField = pFClass.Fields.Field(lngIndex)
    If pField.Editable And pField.Type <> esriFieldTypeGeometry Then
      Set pClone = pField
      strNewName = MyGeneralOperations.ReturnAcceptableFieldName2(pField.Name, pFieldArray, True, False, False, False)
      pSourceNameColl.Add pField.Name, strNewName
      Set pNewField = pClone.Clone
      Set pFieldEdit = pNewField
      With pFieldEdit
        .Name = strNewName
      End With
      pFieldArray.Add pNewField
    End If
  Next lngIndex

  Set pGeoDataset = pFClass
  Set pSpRef = pGeoDataset.SpatialReference

  Dim pNewFClass As IFeatureClass
  Set pNewFClass = MyGeneralOperations.CreateShapefileFeatureClass2(strDir, strFClassName, pSpRef, _
      pFClass.ShapeType, pFieldArray, False)

  Dim lngLinks() As Long
  Dim lngSrcIndex As Long
  Dim lngArrayIndex As Long
  lngArrayIndex = -1
  For lngIndex = 0 To pNewFClass.Fields.FieldCount - 1
    Set pNewField = pNewFClass.Fields.Field(lngIndex)
    If pNewField.Type <> esriFieldTypeGeometry Then
      strNewName = pNewField.Name
      If MyGeneralOperations.CheckCollectionForKey(pSourceNameColl, strNewName) Then
        strSourceName = pSourceNameColl.Item(strNewName)
        lngArrayIndex = lngArrayIndex + 1
        ReDim Preserve lngLinks(1, lngArrayIndex)
        lngLinks(0, lngArrayIndex) = pFClass.FindField(strSourceName)
        lngLinks(1, lngArrayIndex) = pNewFClass.FindField(strNewName)
      End If
    End If
  Next lngIndex

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pNewFCursor As IFeatureCursor
  Dim pNewFBuffer As IFeatureBuffer
  Dim lngCounter As Long

  Set pFCursor = pFClass.Search(Nothing, False)
  Set pNewFCursor = pNewFClass.Insert(True)
  Set pNewFBuffer = pNewFClass.CreateFeatureBuffer

  Dim varVal As Variant

  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    Set pNewFBuffer.Shape = pFeature.ShapeCopy
    For lngIndex = 0 To UBound(lngLinks, 2)
      varVal = pFeature.Value(lngLinks(0, lngIndex))
      If IsNull(varVal) Then
        If pNewFBuffer.Fields.Field(lngLinks(1, lngIndex)).Type = esriFieldTypeString Then
          pNewFBuffer.Value(lngLinks(1, lngIndex)) = ""
        Else
          pNewFBuffer.Value(lngLinks(1, lngIndex)) = -999
        End If
      Else
        pNewFBuffer.Value(lngLinks(1, lngIndex)) = varVal
      End If
    Next lngIndex
    pNewFCursor.InsertFeature pNewFBuffer

    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      pNewFCursor.Flush
    End If

    Set pFeature = pFCursor.NextFeature
  Loop
  pNewFCursor.Flush

  Set CopyFeatureClassToShapefile = pNewFClass

ClearMemory:
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pFieldArray = Nothing
  Set pField = Nothing
  Set pClone = Nothing
  Set pFieldEdit = Nothing
  Set pNewField = Nothing
  Set pSourceNameColl = Nothing
  Set pNewFClass = Nothing
  Erase lngLinks
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pNewFCursor = Nothing
  Set pNewFBuffer = Nothing

End Function

Public Sub ExportFinalDataset_SA()

  Dim lngStart As Long
  lngStart = GetTickCount
  Debug.Print "-----------------------------------"

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFolders As esriSystem.IStringArray

  Dim strOrigRoot As String
  Dim strModRoot As String
  Dim strShiftRoot As String
  Dim strFinalFolder As String
  Call DeclareWorkspaces(strOrigRoot, , strShiftRoot, , strModRoot, , , strFinalFolder)

  Dim pCoverFClass As IFeatureClass
  Dim pDensityFClass As IFeatureClass
  Dim pVegColl As Collection
  Dim pRefColl As Collection
  Set pRefColl = SierraAnchaAnalysis.ReturnSpeciesTypeColl(pCoverFClass, pDensityFClass, pVegColl, , strModRoot)

  Dim strFolder As String
  Dim lngIndex As Long

  Dim pDataset As IDataset
  Dim pGeoDataset As IGeoDataset
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(26912)
  Dim pControlPrecision As IControlPrecision2
  Set pControlPrecision = pSpRef
  Dim pSRRes As ISpatialReferenceResolution
  Set pSRRes = pSpRef
  Dim pSRTol As ISpatialReferenceTolerance
  Set pSRTol = pSpRef
  pSRTol.XYTolerance = 0.0001

  Dim pNewWSFact As IWorkspaceFactory
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Dim pSrcWS As IFeatureWorkspace
  Dim pNewWS As IFeatureWorkspace
  Dim pSrcCoverFClass As IFeatureClass
  Dim pSrcDensFClass As IFeatureClass
  Dim pTopoOp As ITopologicalOperator4
  Dim lngQuadIndex As Long

  Dim strQuadrat As String
  Dim strDestFolder As String
  Dim strItem() As String
  Dim strSite As String
  Dim strSiteSpecific As String
  Dim strPlot As String
  Dim strFileHeader As String
  Dim dblCentroidX As Double
  Dim dblCentroidY As Double

  Dim pDatasetEnum As IEnumDataset
  Dim pWS As IWorkspace

  Dim strFClassName As String
  Dim strNameSplit() As String

  Set pNewWSFact = New FileGDBWorkspaceFactory
  Set pSrcWS = pNewWSFact.OpenFromFile(strShiftRoot & "\Combined_by_Site.gdb", 0)
  Set pNewWS = MyGeneralOperations.CreateOrReturnFileGeodatabase(strFinalFolder & "\Combined_by_Site")

  Set pWS = pSrcWS
  Set pDatasetEnum = pWS.Datasets(esriDTFeatureClass)
  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    Debug.Print strFClassName

    ExportFinalFClass pNewWS, pDataset, pMxDoc, False, pVegColl
    Set pDataset = pDatasetEnum.Next
  Loop

  If Not aml_func_mod.ExistFileDir(strFinalFolder & "\Shapefiles") Then
    MyGeneralOperations.CreateNestedFoldersByPath (strFinalFolder & "\Shapefiles")
  End If
  Set pNewWSFact = New ShapefileWorkspaceFactory
  Set pNewWS = pNewWSFact.OpenFromFile(strFinalFolder & "\Shapefiles", 0)

  pDatasetEnum.Reset

  Set pDataset = pDatasetEnum.Next
  Do Until pDataset Is Nothing
    strFClassName = pDataset.BrowseName
    Debug.Print strFClassName

    ExportFinalFClass pNewWS, pDataset, pMxDoc, True, pVegColl
    Set pDataset = pDatasetEnum.Next
  Loop

  Debug.Print "Done..."

ClearMemory:
  Set pMxDoc = Nothing
  Set pFolders = Nothing
  Set pDataset = Nothing
  Set pGeoDataset = Nothing
  Set pSpRef = Nothing
  Set pControlPrecision = Nothing
  Set pSRRes = Nothing
  Set pSRTol = Nothing
  Set pNewWSFact = Nothing
  Set pSrcWS = Nothing
  Set pNewWS = Nothing
  Set pSrcCoverFClass = Nothing
  Set pSrcDensFClass = Nothing
  Set pTopoOp = Nothing
  Erase strItem
  Set pDatasetEnum = Nothing
  Set pWS = Nothing
  Erase strNameSplit

End Sub


