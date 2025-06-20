Attribute VB_Name = "SierraAncha_Compare"
Option Explicit

Public Sub GenerateRData()

  Dim lngStart As Long
  lngStart = GetTickCount

  Dim pMxDoc As IMxDocument
  Dim pApp As IApplication
  Dim pSBar As IStatusBar
  Dim pProg As IStepProgressor

  Set pMxDoc = ThisDocument
  Set pApp = Application
  Set pSBar = pApp.StatusBar
  Set pProg = pSBar.ProgressBar

  Dim pCoverFClass As IFeatureClass
  Dim pDensityFClass As IFeatureClass
  Dim pVegColl As Collection
  Dim pRefColl As Collection
  Dim strFinalFolder As String

  Set pRefColl = SierraAnchaAnalysis.ReturnSpeciesTypeColl(pCoverFClass, pDensityFClass, pVegColl, , strFinalFolder)

  Dim pParentMaterialColl As Collection
  Set pParentMaterialColl = ReturnParentMaterialPerQuadrat

  Dim strCombinePath As String
  Dim strModifiedRoot As String
  Dim strContainerFolder As String
  Dim strExportPath As String
  Dim datDateUsed As Date

  Call DeclareWorkspaces(strCombinePath, , , , strModifiedRoot, strContainerFolder, , , datDateUsed)
  strExportPath = MyGeneralOperations.MakeUniquedBASEName(strContainerFolder & "\R.Dat." & _
      Format(datDateUsed, "mmddyyyy") & ".csv")

  Dim pWS As IFeatureWorkspace
  Dim pWSFact As IWorkspaceFactory
  Set pWSFact = New FileGDBWorkspaceFactory
  Set pWS = pWSFact.OpenFromFile(strModifiedRoot & "\Combined_by_Site.gdb", 0)

  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim lngDensitySpeciesIndex As Long
  Dim lngDensityQuadratIndex As Long
  Dim lngDensitySiteIndex As Long
  Dim lngDensityYearIndex As Long
  Dim lngDensityAreaIndex As Long
  Dim lngCoverSpeciesIndex As Long
  Dim lngCoverQuadratIndex As Long
  Dim lngCoverSiteIndex As Long
  Dim lngCoverYearIndex As Long
  Dim lngCoverAreaIndex As Long

  Dim pCumulativeAreaColl As New Collection
  Dim dblArea As Double
  Dim dblCumulativeArea As Double

  Set pDensityFClass = pWS.OpenFeatureClass("Density_All")
  lngDensitySpeciesIndex = pDensityFClass.FindField("Species")
  lngDensityQuadratIndex = pDensityFClass.FindField("Quadrat")
  lngDensitySiteIndex = pDensityFClass.FindField("Site")
  lngDensityYearIndex = pDensityFClass.FindField("Year")
  lngDensityAreaIndex = pDensityFClass.FindField("Shape_Area")

  Set pCoverFClass = pWS.OpenFeatureClass("Cover_All")
  lngCoverSpeciesIndex = pCoverFClass.FindField("Species")
  lngCoverQuadratIndex = pCoverFClass.FindField("Quadrat")
  lngCoverSiteIndex = pCoverFClass.FindField("Site")
  lngCoverYearIndex = pCoverFClass.FindField("Year")
  lngCoverAreaIndex = pCoverFClass.FindField("Shape_Area")

  Dim lngCount As Long
  lngCount = pDensityFClass.FeatureCount(Nothing) + pCoverFClass.FeatureCount(Nothing)

  Dim strSpecies As String
  Dim strSite As String
  Dim lngQuadrat As Long
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim lngIndex3 As Long
  Dim lngYear As Long
  Dim strQuadrat As String

  Dim lngAllYears() As Long
  Dim strAllSpecies() As String
  Dim strAllSites() As String
  Dim lngAllQuadrats() As Long
  Dim lngAllSpeciesIndex As Long
  Dim lngAllQuadratsIndex As Long

  Dim pDoneQuadrats As New Collection
  Dim pDoneSpecies As New Collection
  Dim strDrainage As String
  Dim strParentMaterial As String

  lngAllSpeciesIndex = -1
  lngAllQuadratsIndex = -1

  Dim pDrainageColl As New Collection

  pSBar.ShowProgressBar "Initial Pass...", 0, lngCount, 1, True
  pProg.position = 0

  Dim strKey As String
  Dim lngCounter As Long

  lngCounter = 0

  Set pFCursor = pDensityFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))

    If strSpecies <> "" And strSpecies <> "" And strSpecies <> "" Then
      lngYear = CLng(Trim(pFeature.Value(lngDensityYearIndex)))
      strQuadrat = Trim(pFeature.Value(lngDensityQuadratIndex))
      strQuadrat = Replace(strQuadrat, "Q", "", , , vbTextCompare)

      lngQuadrat = CLng(strQuadrat)
      dblArea = pFeature.Value(lngDensityAreaIndex)

      If strSpecies = "Bouteloua curtipendula" And lngYear = 1935 And strQuadrat = "1" Then
        DoEvents
      End If

      MyGeneralOperations.AddValueToLongArray lngYear, lngAllYears, True

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
        pDoneSpecies.Add True, strSpecies
        MyGeneralOperations.AddValueToStringArray strSpecies, strAllSpecies, True
        lngAllSpeciesIndex = UBound(strAllSpecies)
      End If

      strSite = pFeature.Value(lngDensitySiteIndex)
      strDrainage = Replace(strSite, "Natural Drainages Watershed ", "")

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadrats, Format(lngQuadrat, "0")) Then
        pDoneQuadrats.Add True, Format(lngQuadrat, "0")
        MyGeneralOperations.AddValueToLongArray lngQuadrat, lngAllQuadrats, True
      End If

      If Not MyGeneralOperations.CheckCollectionForKey(pDrainageColl, Format(lngQuadrat, "0")) Then
        pDrainageColl.Add strDrainage, Format(lngQuadrat, "0")
      End If

      strKey = strSpecies & "|" & Format(lngYear, "0") & "|" & Format(lngQuadrat, "0")
      If MyGeneralOperations.CheckCollectionForKey(pCumulativeAreaColl, strKey) Then
        dblCumulativeArea = pCumulativeAreaColl.Item(strKey)
        pCumulativeAreaColl.Remove strKey
      Else
        dblCumulativeArea = 0
      End If
      dblCumulativeArea = dblCumulativeArea + dblArea
      pCumulativeAreaColl.Add dblCumulativeArea, strKey
    End If
    Set pFeature = pFCursor.NextFeature
  Loop

  Set pFCursor = pCoverFClass.Search(Nothing, False)
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    pProg.Step
    lngCounter = lngCounter + 1
    If lngCounter Mod 100 = 0 Then
      DoEvents
    End If
    strSpecies = Trim(pFeature.Value(lngDensitySpeciesIndex))

    If strSpecies <> "" And strSpecies <> "" And strSpecies <> "" Then
      lngYear = CLng(Trim(pFeature.Value(lngCoverYearIndex)))
      strQuadrat = Trim(pFeature.Value(lngCoverQuadratIndex))
      strQuadrat = Replace(strQuadrat, "Q", "", , , vbTextCompare)

      lngQuadrat = CLng(strQuadrat)
      dblArea = pFeature.Value(lngCoverAreaIndex)

      If strSpecies = "Bouteloua curtipendula" And lngYear = 1935 And strQuadrat = "1" Then
        DoEvents
      End If

      MyGeneralOperations.AddValueToLongArray lngYear, lngAllYears, True

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneSpecies, strSpecies) Then
        pDoneSpecies.Add True, strSpecies
        MyGeneralOperations.AddValueToStringArray strSpecies, strAllSpecies, True
        lngAllSpeciesIndex = UBound(strAllSpecies)
      End If

      strSite = pFeature.Value(lngCoverSiteIndex)
      strDrainage = Replace(strSite, "Natural Drainages Watershed ", "")

      If Not MyGeneralOperations.CheckCollectionForKey(pDoneQuadrats, Format(lngQuadrat, "0")) Then
        pDoneQuadrats.Add True, Format(lngQuadrat, "0")
        MyGeneralOperations.AddValueToLongArray lngQuadrat, lngAllQuadrats, True
      End If

      If Not MyGeneralOperations.CheckCollectionForKey(pDrainageColl, Format(lngQuadrat, "0")) Then
        pDrainageColl.Add strDrainage, Format(lngQuadrat, "0")
      End If

      strKey = strSpecies & "|" & Format(lngYear, "0") & "|" & Format(lngQuadrat, "0")
      If MyGeneralOperations.CheckCollectionForKey(pCumulativeAreaColl, strKey) Then
        dblCumulativeArea = pCumulativeAreaColl.Item(strKey)
        pCumulativeAreaColl.Remove strKey
      Else
        dblCumulativeArea = 0
      End If
      dblCumulativeArea = dblCumulativeArea + dblArea
      pCumulativeAreaColl.Add dblCumulativeArea, strKey
    End If

    Set pFeature = pFCursor.NextFeature
  Loop

  QuickSort.LongAscending lngAllYears, 0, UBound(lngAllYears)
  QuickSort.LongAscending lngAllQuadrats, 0, UBound(lngAllQuadrats)
  QuickSort.StringsAscending strAllSpecies, 0, UBound(strAllSpecies)
  Dim strYear As String
  Dim varVegArray() As Variant
  Dim strFuncType As String
  Dim strSpeciesCode As String
  Dim strLifeCycle As String
  Dim strC3C4 As String

  lngCount = (UBound(lngAllYears) + 1) * (UBound(lngAllQuadrats) + 1) * (UBound(strAllSpecies) + 1)
  lngCounter = 0
  Dim strSubReport As String
  strSubReport = """site"",""year"",""drainage"",""parent.material"",""quadrat"",""species"",""photosynthetic.pathway"",""functional.type"",""life.cycle"",""p.cover"""
  MyGeneralOperations.WriteTextFile strExportPath, strSubReport, False, False

  strSubReport = ""

  pSBar.ShowProgressBar "Writing Data...", 0, lngCount, 1, True
  pProg.position = 0

  strSite = "ND"

  For lngIndex = 0 To UBound(lngAllYears)
    lngYear = lngAllYears(lngIndex)
    If lngYear = 1940 Or lngYear = 1941 Then
      strYear = "1940-41"
    Else
      strYear = Format(lngYear, "0")
    End If

    For lngIndex2 = 0 To UBound(lngAllQuadrats)
      lngQuadrat = lngAllQuadrats(lngIndex2)
      strParentMaterial = pParentMaterialColl.Item(Format(lngQuadrat, "0"))
      strDrainage = pDrainageColl.Item(Format(lngQuadrat, "0"))

      For lngIndex3 = 0 To UBound(strAllSpecies)
        strSpecies = strAllSpecies(lngIndex3)
        varVegArray = pVegColl.Item(strSpecies)
        strFuncType = varVegArray(8)
        strSpeciesCode = varVegArray(9)
        strLifeCycle = varVegArray(5)
        strC3C4 = varVegArray(6)

        If strSpeciesCode = "" Then
          DoEvents
        End If

        If strSpecies <> "No Cover Species Observed" And strSpecies <> "No Density Species Observed" Then
          strKey = strSpecies & "|" & Format(lngYear, "0") & "|" & Format(lngQuadrat, "0")
          If MyGeneralOperations.CheckCollectionForKey(pCumulativeAreaColl, strKey) Then
            dblCumulativeArea = pCumulativeAreaColl.Item(strKey)
          Else
            dblCumulativeArea = 0
          End If

          strSubReport = strSubReport & strSite & "," & strYear & "," & strDrainage & "," & strParentMaterial & "," & Format(lngQuadrat, "0") & _
              "," & strSpeciesCode & "," & strC3C4 & "," & strFuncType & "," & strLifeCycle & "," & Format(dblCumulativeArea * 100, "0.00000000") & vbCrLf

          lngCounter = lngCounter + 1
          pProg.Step
          If lngCounter = 100 Then
            DoEvents
            strSubReport = Left(strSubReport, Len(strSubReport) - 2)
            MyGeneralOperations.WriteTextFile strExportPath, strSubReport, False, True
            lngCounter = 0
            strSubReport = ""
          End If
        End If
      Next lngIndex3
    Next lngIndex2
  Next lngIndex

  MyGeneralOperations.WriteTextFile strExportPath, strSubReport, False, True
  Debug.Print "Done..."

  pSBar.HideProgressBar
  pProg.position = 0
  Debug.Print "Done..."
  Debug.Print MyGeneralOperations.ReturnTimeElapsedFromMilliseconds(GetTickCount - lngStart)

ClearMemory:
  Set pMxDoc = Nothing
  Set pApp = Nothing
  Set pSBar = Nothing
  Set pProg = Nothing
  Set pWS = Nothing
  Set pWSFact = Nothing
  Set pDensityFClass = Nothing
  Set pCoverFClass = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Erase strAllSpecies
  Erase strAllSites
  Erase lngAllQuadrats
  Set pDoneSpecies = Nothing

End Sub


